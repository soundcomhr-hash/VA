// ---------- Settings (endpoint + token), saved locally on the phone ----------
const settings = {
  get endpoint() { return localStorage.getItem('va_endpoint') || ''; },
  set endpoint(v) { localStorage.setItem('va_endpoint', v); },
  get token() { return localStorage.getItem('va_token') || ''; },
  set token(v) { localStorage.setItem('va_token', v); },
};

// ---------- IndexedDB offline queue ----------
const DB_NAME = 'va-queue';
const STORE = 'items';

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) {
        db.createObjectStore(STORE, { keyPath: 'id', autoIncrement: true });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function queueAdd(item) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    const req = tx.objectStore(STORE).add(item);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function queueAll() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readonly');
    const req = tx.objectStore(STORE).getAll();
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function queueDelete(id) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    const req = tx.objectStore(STORE).delete(id);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

async function refreshQueueBadge() {
  const items = await queueAll();
  const wrap = document.getElementById('queueBadgeWrap');
  const badge = document.getElementById('queueBadge');
  badge.textContent = String(items.length);
  wrap.hidden = items.length === 0;
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

// One-tap setup: opening the app with ?endpoint=...&token=... saves the
// server settings automatically (so Ziv never types them on the phone).
(function applySetupLink() {
  const params = new URLSearchParams(location.search);
  if (params.get('endpoint')) {
    settings.endpoint = params.get('endpoint');
    if (params.get('token')) settings.token = params.get('token');
    history.replaceState(null, '', location.pathname);
  }
})();

// ---------- Backend API ----------
// Body goes as text/plain: Apps Script web apps can't answer the browser's
// CORS "preflight" check that a JSON content-type would trigger.
async function apiPost(action, payload) {
  const res = await fetch(settings.endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(Object.assign({ action, token: settings.token }, payload || {})),
  });
  if (!res.ok) throw new Error('bad status ' + res.status);
  return res.json();
}

// ---------- Stale-while-revalidate cache ----------
// Each screen shows its last-known data instantly from localStorage, then
// refreshes from the server in the background. This is the standard pattern
// (SWR / React Query) for hiding a slow backend - the user never waits on a
// blank "loading" screen when we already have something to show.
function cacheGet(key) {
  try { return JSON.parse(localStorage.getItem('va_cache_' + key)); }
  catch (e) { return null; }
}
function cacheSet(key, val) {
  try { localStorage.setItem('va_cache_' + key, JSON.stringify(val)); }
  catch (e) {}
}

// ---------- Sync queued items to the backend ----------
let syncing = false;

async function attemptSync() {
  if (syncing) return;
  if (!navigator.onLine) return;
  if (!settings.endpoint) {
    setSendNote('לא הוגדרה עדיין כתובת שרת — ההקלטות שמורות במכשיר וימתינו.');
    return;
  }
  syncing = true;
  try {
    const items = await queueAll();
    for (const item of items) {
      try {
        const audioBase64 = item.audioBlob ? await blobToBase64(item.audioBlob) : null;
        const result = await apiPost('capture', {
          text: item.text,
          audioBase64,
          audioMime: item.mime || null,
          needsTranscription: !!item.needsTranscription,
          createdAt: item.createdAt,
        });
        if (!result.ok) throw new Error(result.error || 'server-error');
        await queueDelete(item.id);
        setSendNote(result.summary
          ? 'נקלט: ' + result.summary + (result.question ? ' · ממתין לתשובה שלך 📥' : '')
          : 'נשלח לשרת בהצלחה');
      } catch (err) {
        // Network/server not ready yet - leave item queued and stop this pass.
        break;
      }
    }
  } finally {
    syncing = false;
    refreshQueueBadge();
  }
}

// ---------- Tabs ----------
function setupTabs() {
  const buttons = document.querySelectorAll('.tab-btn');
  buttons.forEach((btn) => {
    btn.addEventListener('click', () => {
      buttons.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      document.querySelectorAll('.tab-panel').forEach((p) => (p.hidden = true));
      document.getElementById('tab-' + btn.dataset.tab).hidden = false;
      if (btn.dataset.tab === 'inbox') loadInbox();
      if (btn.dataset.tab === 'today') loadToday();
      if (btn.dataset.tab === 'tasks') loadTasks();
    });
  });
}

// ---------- משימות (tasks) screen ----------
function setTasksStatus(msg) {
  document.getElementById('tasksStatus').textContent = msg;
}

// renderedTasks is the array currently on screen - the single source of truth
// for optimistic edits (toggle / color change mutate it and re-render at once).
let renderedTasks = [];

async function loadTasks() {
  if (!settings.endpoint) {
    document.getElementById('tasksList').innerHTML = '';
    setTasksStatus('עדיין לא הוגדר שרת (⚙ למעלה).');
    return;
  }
  const cached = cacheGet('tasks');
  if (cached) renderTasks(cached);        // instant, no waiting
  else setTasksStatus('טוען...');
  if (!navigator.onLine) {
    if (!cached) setTasksStatus('אין חיבור לאינטרנט - נסו שוב כשיש קליטה.');
    return;
  }
  try {
    const result = await apiPost('tasks_list');
    if (!result.ok) throw new Error(result.error);
    cacheSet('tasks', result.items);
    renderTasks(result.items);
  } catch (err) {
    if (!cached) setTasksStatus('שגיאה בטעינה מהשרת. נסו לרענן.');
  }
}

// Optimistic: flip the UI instantly, save in the background, revert on failure.
function toggleTaskOptimistic(item) {
  const prev = item.status;
  item.status = (prev === 'בוצע') ? 'פתוח' : 'בוצע';
  renderTasks(renderedTasks);
  apiPost('tasks_toggle', { id: item.id })
    .then((r) => { if (!r || !r.ok) throw new Error(); cacheSet('tasks', renderedTasks); })
    .catch(() => {
      item.status = prev;
      renderTasks(renderedTasks);
      setTasksStatus('לא נשמר - בעיית חיבור. נסו שוב.');
    });
}

function setTaskColorOptimistic(item, color) {
  const prev = item.color;
  if (prev === color) return;
  item.color = color;
  renderTasks(renderedTasks);
  apiPost('tasks_set_color', { id: item.id, color })
    .then((r) => { if (!r || !r.ok) throw new Error(); cacheSet('tasks', renderedTasks); })
    .catch(() => {
      item.color = prev;
      renderTasks(renderedTasks);
      setTasksStatus('לא נשמר - בעיית חיבור. נסו שוב.');
    });
}

const TASK_COLORS = ['אדום', 'כחול', 'ירוק', 'צהוב'];

const CONGRATS_MESSAGES = [
  'כל הכבוד! הכל בוצע 🎉',
  'סיימת הכל, יפה מאוד 💪',
  'רשימה נקייה - עבודה טובה 👏',
  'הכל מסומן. תהנה מהיום 🌿',
  'זהו זה, אין מה להוסיף 🙌',
];
function pickCongrats() {
  return CONGRATS_MESSAGES[Math.floor(Math.random() * CONGRATS_MESSAGES.length)];
}

function renderTasks(items) {
  renderedTasks = items;
  const list = document.getElementById('tasksList');
  list.innerHTML = '';
  if (!items.length) {
    setTasksStatus('אין משימות פתוחות 🎉');
    return;
  }
  setTasksStatus(items.every((i) => i.status === 'בוצע') ? pickCongrats() : '');
  items.forEach((item) => {
    const card = document.createElement('div');
    card.className = 'task-card color-' + item.color +
      (item.status === 'בוצע' ? ' done' : '') +
      (item.urgent && item.status !== 'בוצע' ? ' urgent' : '');

    const check = document.createElement('button');
    check.className = 'task-check';
    check.textContent = item.status === 'בוצע' ? '✓' : '';
    check.addEventListener('click', () => toggleTaskOptimistic(item));
    card.appendChild(check);

    const text = document.createElement('div');
    text.className = 'task-text';
    text.textContent = (item.urgent && item.status !== 'בוצע' ? '⏰ ' : '') +
      item.text + (item.dueTime ? ' · ' + item.dueTime : '');
    card.appendChild(text);

    const colors = document.createElement('div');
    colors.className = 'task-colors';
    TASK_COLORS.forEach((color) => {
      const dot = document.createElement('button');
      dot.className = 'color-' + color + (color === item.color ? ' selected' : '');
      dot.title = color;
      dot.addEventListener('click', () => setTaskColorOptimistic(item, color));
      colors.appendChild(dot);
    });
    card.appendChild(colors);

    list.appendChild(card);
  });
}

function setupMorningButtons() {
  document.getElementById('morningBtn').addEventListener('click', async () => {
    if (!settings.endpoint) { setTasksStatus('עדיין לא הוגדר שרת (⚙ למעלה).'); return; }
    setTasksStatus('טוען...');
    try {
      const result = await apiPost('morning_start');
      if (!result.ok) throw new Error(result.error);
      renderTasks(result.items);
      setTasksStatus(result.items.length
        ? 'בוקר טוב! מציג רק את משימות היום - "רענן" מציג הכל'
        : 'בוקר טוב! אין משימות להיום 🎉');
    } catch (e) {
      setTasksStatus('שגיאה. נסו שוב.');
    }
  });

  document.getElementById('snoozeBtn').addEventListener('click', async () => {
    try {
      const r = await apiPost('morning_snooze');
      if (!r.ok) throw new Error(r.error);
      setTasksStatus('התזכורות הושתקו עד מחר.');
    } catch (e) {
      setTasksStatus('שגיאה. נסו שוב.');
    }
  });
}

// ---------- היום (today) screen ----------
function setTodayStatus(msg) {
  document.getElementById('todayStatus').textContent = msg;
}

async function loadToday() {
  document.getElementById('todayTitle').textContent =
    new Date().toLocaleDateString('he-IL', { weekday: 'long', day: 'numeric', month: 'long' });
  if (!settings.endpoint) {
    document.getElementById('todayList').innerHTML = '';
    setTodayStatus('עדיין לא הוגדר שרת (⚙ למעלה).');
    return;
  }
  const cached = cacheGet('today');
  if (cached) {
    renderToday(cached.events || []);
    renderTodayPending(cached.pending || []);
  } else {
    setTodayStatus('טוען את הלו"ז...');
  }
  if (!navigator.onLine) {
    if (!cached) setTodayStatus('אין חיבור לאינטרנט - נסו שוב כשיש קליטה.');
    return;
  }
  try {
    const result = await apiPost('today');
    if (!result.ok) throw new Error(result.error);
    const data = { events: result.events || [], pending: result.pending || [] };
    cacheSet('today', data);
    renderToday(data.events);
    renderTodayPending(data.pending);
  } catch (err) {
    if (!cached) setTodayStatus('שגיאה בטעינה מהיומן. נסו לרענן.');
  }
}

function renderTodayPending(items) {
  const wrap = document.getElementById('todayPendingWrap');
  const list = document.getElementById('todayPendingList');
  list.innerHTML = '';
  wrap.hidden = items.length === 0;
  items.forEach((item) => {
    const card = document.createElement('div');
    card.className = 'today-pending-card';

    const body = document.createElement('div');
    body.className = 'pending-body';
    body.textContent = item.summary;
    if (item.question) {
      const q = document.createElement('div');
      q.className = 'pending-question';
      q.textContent = '❓ ' + item.question;
      body.appendChild(q);
    }
    card.appendChild(body);

    const goBtn = document.createElement('button');
    goBtn.className = 'btn btn-secondary';
    goBtn.textContent = 'לענות';
    goBtn.addEventListener('click', () => {
      document.querySelector('.tab-btn[data-tab="inbox"]').click();
    });
    card.appendChild(goBtn);

    list.appendChild(card);
  });
}

function renderToday(events) {
  const list = document.getElementById('todayList');
  list.innerHTML = '';
  if (!events.length) {
    setTodayStatus('אין אירועים ביומן היום 🎉');
    return;
  }
  setTodayStatus('');
  const nowMinutes = new Date().getHours() * 60 + new Date().getMinutes();
  const toMinutes = (hhmm) => parseInt(hhmm.slice(0, 2), 10) * 60 + parseInt(hhmm.slice(3), 10);

  events.forEach((ev) => {
    const card = document.createElement('div');
    card.className = 'today-event';
    if (!ev.allDay) {
      if (toMinutes(ev.end) < nowMinutes) card.classList.add('past');
      else if (toMinutes(ev.start) <= nowMinutes) card.classList.add('current');
    }

    const time = document.createElement('div');
    time.className = 'event-time';
    time.textContent = ev.allDay ? 'כל היום' : ev.start + '–' + ev.end;
    card.appendChild(time);

    const body = document.createElement('div');
    body.className = 'event-body';
    const title = document.createElement('div');
    title.className = 'event-title';
    title.textContent = ev.title || '(ללא כותרת)';
    body.appendChild(title);
    if (ev.location) {
      const loc = document.createElement('div');
      loc.className = 'event-location';
      loc.textContent = '📍 ' + ev.location;
      body.appendChild(loc);
    }
    card.appendChild(body);
    list.appendChild(card);
  });
}

// ---------- ממתין (inbox) screen ----------
function setInboxStatus(msg) {
  document.getElementById('inboxStatus').textContent = msg;
}

async function loadInbox() {
  if (!settings.endpoint) {
    document.getElementById('inboxList').innerHTML = '';
    setInboxStatus('עדיין לא הוגדר שרת (⚙ למעלה). כשהשרת יחובר, הפריטים יופיעו כאן.');
    return;
  }
  const cached = cacheGet('inbox');
  if (cached) renderInbox(cached);
  else setInboxStatus('טוען...');
  if (!navigator.onLine) {
    if (!cached) setInboxStatus('אין חיבור לאינטרנט - נסו שוב כשיש קליטה.');
    return;
  }
  try {
    const result = await apiPost('inbox_list');
    if (!result.ok) throw new Error(result.error);
    cacheSet('inbox', result.items);
    renderInbox(result.items);
  } catch (err) {
    if (!cached) setInboxStatus('שגיאה בטעינה מהשרת. נסו לרענן.');
  }
}

// Drop a resolved card from the list instantly (after confirm/delete) instead
// of waiting on a full server reload.
function removeInboxLocally(id) {
  const remaining = (cacheGet('inbox') || []).filter((it) => it.id !== id);
  cacheSet('inbox', remaining);
  renderInbox(remaining);
}

function renderInbox(items) {
  const badge = document.getElementById('inboxTabBadge');
  badge.textContent = String(items.length);
  badge.hidden = items.length === 0;
  const list = document.getElementById('inboxList');
  list.innerHTML = '';
  if (!items.length) {
    setInboxStatus('אין פריטים ממתינים 🎉');
    return;
  }
  setInboxStatus('');
  items.forEach((item) => {
    const card = document.createElement('div');
    card.className = 'inbox-card';

    const kind = document.createElement('div');
    kind.className = 'card-kind';
    kind.textContent = item.kind + ' · ' + item.created;
    card.appendChild(kind);

    const summary = document.createElement('div');
    summary.className = 'card-summary';
    summary.textContent = item.summary;
    card.appendChild(summary);

    if (item.audioUrl) {
      const audio = document.createElement('div');
      audio.className = 'card-audio';
      const link = document.createElement('a');
      link.href = item.audioUrl;
      link.target = '_blank';
      link.rel = 'noopener';
      link.textContent = '🎧 האזנה להקלטה';
      audio.appendChild(link);
      card.appendChild(audio);
    }

    if (item.question) {
      const q = document.createElement('div');
      q.className = 'card-question';
      q.textContent = '❓ ' + item.question;
      card.appendChild(q);

      const row = document.createElement('div');
      row.className = 'answer-row';
      const input = document.createElement('input');
      input.type = 'text';
      input.placeholder = 'תשובה...';
      const btn = document.createElement('button');
      btn.className = 'btn btn-primary';
      btn.textContent = 'ענה';
      btn.addEventListener('click', async () => {
        if (!input.value.trim()) return;
        btn.disabled = true;
        try {
          const r = await apiPost('inbox_answer', { id: item.id, answer: input.value.trim() });
          if (!r.ok) throw new Error(r.error);
          await loadInbox();
        } catch (e) {
          setInboxStatus('שגיאה בשליחת התשובה. נסו שוב.');
          btn.disabled = false;
        }
      });
      input.addEventListener('keydown', (e) => { if (e.key === 'Enter') btn.click(); });
      row.appendChild(input);
      row.appendChild(btn);
      card.appendChild(row);

      if (item.question.indexOf('איפה?') !== -1) {
        const noPlaceBtn = document.createElement('button');
        noPlaceBtn.className = 'btn btn-noplace';
        noPlaceBtn.textContent = 'אין מקום';
        noPlaceBtn.addEventListener('click', async () => {
          noPlaceBtn.disabled = true;
          try {
            const r = await apiPost('inbox_answer', { id: item.id, answer: '', noPlace: true });
            if (!r.ok) throw new Error(r.error);
            await loadInbox();
          } catch (e) {
            setInboxStatus('שגיאה. נסו שוב.');
            noPlaceBtn.disabled = false;
          }
        });
        card.appendChild(noPlaceBtn);
      }
    }

    const actions = document.createElement('div');
    actions.className = 'card-actions';

    const confirmBtn = document.createElement('button');
    confirmBtn.className = 'btn btn-confirm';
    confirmBtn.textContent = 'אשר ✓';
    confirmBtn.addEventListener('click', async () => {
      confirmBtn.disabled = true;
      try {
        const r = await apiPost('inbox_confirm', { id: item.id });
        if (!r.ok) {
          setInboxStatus(r.question
            ? (r.error === 'incomplete' ? 'אי אפשר לאשר - חסר מידע: ' : '') + r.question
            : 'שגיאה באישור.');
          confirmBtn.disabled = false;
          return;
        }
        setInboxStatus(r.message || 'אושר ✓');
        removeInboxLocally(item.id);
      } catch (e) {
        setInboxStatus('שגיאה באישור. נסו שוב.');
        confirmBtn.disabled = false;
      }
    });
    actions.appendChild(confirmBtn);

    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'btn btn-danger';
    deleteBtn.textContent = 'מחק';
    deleteBtn.addEventListener('click', async () => {
      if (!confirm('למחוק את הפריט? (הטקסט המקורי יישאר רשום בגיליון "קלט")')) return;
      deleteBtn.disabled = true;
      try {
        const r = await apiPost('inbox_delete', { id: item.id });
        if (!r.ok) throw new Error(r.error);
        removeInboxLocally(item.id);
      } catch (e) {
        setInboxStatus('שגיאה במחיקה. נסו שוב.');
        deleteBtn.disabled = false;
      }
    });
    actions.appendChild(deleteBtn);

    card.appendChild(actions);
    list.appendChild(card);
  });
}

// ---------- Settings modal ----------
function setupSettings() {
  const modal = document.getElementById('settingsModal');
  const endpointInput = document.getElementById('endpointInput');
  const tokenInput = document.getElementById('tokenInput');

  document.getElementById('settingsBtn').addEventListener('click', () => {
    endpointInput.value = settings.endpoint;
    tokenInput.value = settings.token;
    modal.hidden = false;
  });
  document.getElementById('settingsClose').addEventListener('click', () => {
    modal.hidden = true;
  });
  document.getElementById('settingsSave').addEventListener('click', () => {
    settings.endpoint = endpointInput.value.trim();
    settings.token = tokenInput.value.trim();
    modal.hidden = true;
    attemptSync();
  });
}

// ---------- Connection status ----------
function setupConnStatus() {
  const el = document.getElementById('connStatus');
  function update() {
    if (navigator.onLine) {
      el.classList.remove('offline');
      el.title = 'מחובר';
    } else {
      el.classList.add('offline');
      el.title = 'לא מחובר - עובד במצב לא מקוון';
    }
  }
  window.addEventListener('online', () => { update(); attemptSync(); });
  window.addEventListener('offline', update);
  update();
}

function setSendNote(msg) {
  document.getElementById('sendNote').textContent = msg;
}

// ---------- Recording ----------
// Android cannot feed the microphone to SpeechRecognition and MediaRecorder
// at the same time (recognition hears garbled audio and returns nomatch), so
// capture is one of two exclusive modes per Ziv's decision:
//   - big mic button  = live transcription only (no audio saved)
//   - small audio button = audio recording only, queued as "לתמלול"
const SpeechRecognitionImpl = window.SpeechRecognition || window.webkitSpeechRecognition;

let recognition = null;
let mediaRecorder = null;
let audioChunks = [];
let currentAudioBlob = null;
let currentAudioMime = null;
let activeMode = null; // null | 'transcribe' | 'audio'
let finalTranscript = '';
let recognitionFatal = false;

function setupRecording() {
  const micBtn = document.getElementById('micBtn');
  const audioBtn = document.getElementById('audioBtn');
  const micStatus = document.getElementById('micStatus');
  const transcriptBox = document.getElementById('transcript');

  if (!SpeechRecognitionImpl) {
    micBtn.classList.add('disabled');
    micBtn.disabled = true;
    micStatus.textContent = 'אין זיהוי דיבור בדפדפן הזה - השתמשו בהקלטת אודיו או הקלידו';
  }

  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    audioBtn.disabled = true;
    audioBtn.textContent = 'הקלטת אודיו לא נתמכת בדפדפן הזה';
  }

  micBtn.addEventListener('click', () => {
    if (activeMode === 'transcribe') {
      stopAll();
    } else if (activeMode === null) {
      startTranscription();
    }
  });

  audioBtn.addEventListener('click', async () => {
    if (activeMode === 'audio') {
      stopAll();
    } else if (activeMode === null) {
      await startAudioRecording();
    }
  });

  function startTranscription() {
    finalTranscript = transcriptBox.value ? transcriptBox.value + ' ' : '';
    recognition = new SpeechRecognitionImpl();
    recognition.lang = 'he-IL';
    // Single-utterance mode: continuous/interim is unreliable on Android
    // (fires audio/speech events but never delivers results). onend
    // auto-restarts so the session feels continuous anyway.
    recognition.continuous = false;
    recognition.interimResults = false;
    recognition.onresult = (event) => {
      for (let i = event.resultIndex; i < event.results.length; i++) {
        if (event.results[i].isFinal) {
          finalTranscript += event.results[i][0].transcript + ' ';
        }
      }
      transcriptBox.value = finalTranscript.trim();
    };
    recognition.onerror = (event) => {
      const messages = {
        'not-allowed': 'אין הרשאה למיקרופון - צריך לאשר בדפדפן',
        'service-not-allowed': 'זיהוי דיבור חסום בדפדפן הזה - השתמשו בהקלטת אודיו',
        'network': 'אין אינטרנט לזיהוי דיבור - השתמשו בהקלטת אודיו בלבד',
        'audio-capture': 'בעיה בגישה למיקרופון',
        'no-speech': null,
        'aborted': null,
      };
      const msg = messages.hasOwnProperty(event.error) ? messages[event.error] : ('שגיאת זיהוי דיבור: ' + event.error);
      if (msg) micStatus.textContent = msg;
      if (event.error === 'not-allowed' || event.error === 'service-not-allowed' ||
          event.error === 'audio-capture' || event.error === 'network') {
        recognitionFatal = true;
      }
    };
    recognition.onend = () => {
      if (activeMode === 'transcribe' && !recognitionFatal) {
        try { recognition.start(); } catch (e) {}
      } else if (recognitionFatal) {
        stopAll(true);
      }
    };
    recognitionFatal = false;
    try { recognition.start(); } catch (e) {
      micStatus.textContent = 'זיהוי הדיבור לא נדלק - נסו שוב';
      return;
    }
    activeMode = 'transcribe';
    micBtn.classList.add('recording');
    micStatus.textContent = 'מקשיב... דברו, ואחרי משפט עצרו שנייה. לחצו לסיום';
  }

  async function startAudioRecording() {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      audioChunks = [];
      currentAudioBlob = null;
      currentAudioMime = 'audio/webm';
      mediaRecorder = new MediaRecorder(stream);
      mediaRecorder.ondataavailable = (e) => { if (e.data.size > 0) audioChunks.push(e.data); };
      mediaRecorder.onstop = () => {
        currentAudioBlob = new Blob(audioChunks, { type: currentAudioMime });
        stream.getTracks().forEach((t) => t.stop());
        setSendNote('האודיו מוכן - לחצו שלח והוא יסומן "לתמלול"');
      };
      mediaRecorder.start();
    } catch (err) {
      micStatus.textContent = 'לא ניתן לגשת למיקרופון - צריך לאשר הרשאה';
      return;
    }
    activeMode = 'audio';
    audioBtn.classList.add('recording');
    audioBtn.textContent = '⏹ עצור הקלטת אודיו';
    micStatus.textContent = 'מקליט אודיו (ללא תמלול חי)...';
  }

  function stopAll(keepStatus) {
    activeMode = null;
    micBtn.classList.remove('recording');
    audioBtn.classList.remove('recording');
    audioBtn.textContent = '🔴 הקלטת אודיו בלבד (לתמלול מאוחר)';
    if (!keepStatus) micStatus.textContent = 'לחצו כדי לדבר';
    if (recognition) {
      try { recognition.stop(); } catch (e) {}
      recognition = null;
    }
    if (mediaRecorder && mediaRecorder.state !== 'inactive') {
      mediaRecorder.stop();
    }
  }
}

function resetCapture() {
  document.getElementById('transcript').value = '';
  currentAudioBlob = null;
  finalTranscript = '';
}

// ---------- Send / Discard ----------
function setupSendDiscard() {
  document.getElementById('sendBtn').addEventListener('click', async () => {
    const text = document.getElementById('transcript').value.trim();
    if (!text && !currentAudioBlob) {
      setSendNote('אין מה לשלוח - דברו או הקלידו טקסט קודם.');
      return;
    }
    await queueAdd({
      text,
      audioBlob: currentAudioBlob,
      mime: currentAudioMime,
      needsTranscription: !!currentAudioBlob && !text,
      createdAt: new Date().toISOString(),
    });
    resetCapture();
    await refreshQueueBadge();
    setSendNote('נשמר. מנסה לשלוח לשרת...');
    attemptSync();
  });

  document.getElementById('discardBtn').addEventListener('click', () => {
    resetCapture();
    setSendNote('');
  });
}

// ---------- Install prompt ----------
function setupInstall() {
  let deferredPrompt = null;
  const installBtn = document.getElementById('installBtn');
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    installBtn.hidden = false;
  });
  installBtn.addEventListener('click', async () => {
    installBtn.hidden = true;
    if (deferredPrompt) {
      deferredPrompt.prompt();
      await deferredPrompt.userChoice;
      deferredPrompt = null;
    }
  });
  window.addEventListener('appinstalled', () => { installBtn.hidden = true; });
}

// ---------- Service worker ----------
function registerServiceWorker() {
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
}

// ---------- Init ----------
document.getElementById('inboxRefresh').addEventListener('click', loadInbox);
document.getElementById('todayRefresh').addEventListener('click', loadToday);
document.getElementById('tasksRefresh').addEventListener('click', loadTasks);
setupMorningButtons();

setupTabs();
if (settings.endpoint && navigator.onLine) loadInbox();
setupSettings();
setupConnStatus();
setupRecording();
setupSendDiscard();
setupInstall();
registerServiceWorker();
refreshQueueBadge();
attemptSync();
