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
        setSendNote(result.summary ? 'נקלט: ' + result.summary : 'נשלח לשרת בהצלחה');
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
    });
  });
}

// ---------- ממתין (inbox) screen ----------
function setInboxStatus(msg) {
  document.getElementById('inboxStatus').textContent = msg;
}

async function loadInbox() {
  const list = document.getElementById('inboxList');
  if (!settings.endpoint) {
    list.innerHTML = '';
    setInboxStatus('עדיין לא הוגדר שרת (⚙ למעלה). כשהשרת יחובר, הפריטים יופיעו כאן.');
    return;
  }
  if (!navigator.onLine) {
    setInboxStatus('אין חיבור לאינטרנט - נסו שוב כשיש קליטה.');
    return;
  }
  setInboxStatus('טוען...');
  try {
    const result = await apiPost('inbox_list');
    if (!result.ok) throw new Error(result.error);
    renderInbox(result.items);
  } catch (err) {
    setInboxStatus('שגיאה בטעינה מהשרת. נסו לרענן.');
  }
}

function renderInbox(items) {
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
      row.appendChild(input);
      row.appendChild(btn);
      card.appendChild(row);
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
          setInboxStatus(r.error === 'incomplete'
            ? 'אי אפשר לאשר - חסר מידע: ' + (r.question || '')
            : 'שגיאה באישור.');
          confirmBtn.disabled = false;
          return;
        }
        setInboxStatus(r.message || 'אושר ✓');
        await loadInbox();
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
        await loadInbox();
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

setupTabs();
setupSettings();
setupConnStatus();
setupRecording();
setupSendDiscard();
setupInstall();
registerServiceWorker();
refreshQueueBadge();
attemptSync();
