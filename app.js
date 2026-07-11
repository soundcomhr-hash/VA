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
        const res = await fetch(settings.endpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'capture',
            token: settings.token,
            text: item.text,
            audioBase64,
            audioMime: item.mime || null,
            createdAt: item.createdAt,
          }),
        });
        if (!res.ok) throw new Error('bad status ' + res.status);
        await queueDelete(item.id);
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
    });
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

// ---------- Recording (speech recognition + audio capture) ----------
const SpeechRecognitionImpl = window.SpeechRecognition || window.webkitSpeechRecognition;

let recognition = null;
let mediaRecorder = null;
let audioChunks = [];
let currentAudioBlob = null;
let currentAudioMime = null;
let isRecording = false;
let finalTranscript = '';
let recognitionFatal = false;

// Temporary diagnostic trail shown on screen (remove once speech works on device)
const debugEvents = [];
function debugLog(evt) {
  debugEvents.push(evt);
  if (debugEvents.length > 6) debugEvents.shift();
  const el = document.getElementById('debugLine');
  if (el) el.textContent = debugEvents.join(' | ');
}

function setupRecording() {
  const micBtn = document.getElementById('micBtn');
  const micStatus = document.getElementById('micStatus');
  const transcriptBox = document.getElementById('transcript');

  if (!SpeechRecognitionImpl) {
    micStatus.textContent = 'זיהוי דיבור לא זמין בדפדפן הזה - האודיו יוקלט, אפשר להקליד ידנית';
  }

  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    micBtn.classList.add('disabled');
    micBtn.disabled = true;
    micStatus.textContent = 'הקלטת אודיו לא נתמכת בדפדפן הזה';
    return;
  }

  micBtn.addEventListener('click', async () => {
    if (isRecording) {
      stopRecording();
    } else {
      await startRecording();
    }
  });

  // TEST BUILD (v5): audio recording temporarily disabled to confirm the
  // Android mic contention between MediaRecorder and SpeechRecognition.
  const RECORD_AUDIO = false;

  async function startRecording() {
    if (RECORD_AUDIO) {
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
        };
        mediaRecorder.start();
      } catch (err) {
        micStatus.textContent = 'לא ניתן לגשת למיקרופון - צריך לאשר הרשאה';
        return;
      }
    }

    finalTranscript = transcriptBox.value ? transcriptBox.value + ' ' : '';
    if (!SpeechRecognitionImpl) debugLog('no-speech-api');
    if (SpeechRecognitionImpl) {
      recognition = new SpeechRecognitionImpl();
      recognition.lang = 'he-IL';
      recognition.continuous = true;
      recognition.interimResults = true;
      recognition.onstart = () => debugLog('start');
      recognition.onaudiostart = () => debugLog('audio');
      recognition.onspeechstart = () => debugLog('speech');
      recognition.onresult = (event) => {
        debugLog('result');
        let interim = '';
        for (let i = event.resultIndex; i < event.results.length; i++) {
          const chunk = event.results[i][0].transcript;
          if (event.results[i].isFinal) {
            finalTranscript += chunk + ' ';
          } else {
            interim += chunk;
          }
        }
        transcriptBox.value = (finalTranscript + interim).trim();
      };
      recognition.onerror = (event) => {
        debugLog('err:' + event.error);
        const messages = {
          'not-allowed': 'אין הרשאה לזיהוי דיבור - אפשר לדבר, האודיו עדיין מוקלט',
          'service-not-allowed': 'זיהוי דיבור חסום בדפדפן - האודיו עדיין מוקלט',
          'network': 'אין אינטרנט לזיהוי דיבור - האודיו עדיין מוקלט, יתמלל אחר כך',
          'audio-capture': 'בעיה בגישה למיקרופון עבור זיהוי דיבור',
          'no-speech': null,
          'aborted': null,
        };
        const msg = messages.hasOwnProperty(event.error) ? messages[event.error] : ('שגיאת זיהוי דיבור: ' + event.error);
        if (msg) micStatus.textContent = msg;
        if (event.error === 'not-allowed' || event.error === 'service-not-allowed' || event.error === 'audio-capture') {
          recognitionFatal = true;
        }
      };
      recognition.onend = () => {
        debugLog('end');
        if (isRecording && !recognitionFatal) {
          try { recognition.start(); } catch (e) {}
        }
      };
      recognitionFatal = false;
      try { recognition.start(); } catch (e) {}
    }

    isRecording = true;
    micBtn.classList.add('recording');
    micStatus.textContent = 'מקליט... לחצו כדי לעצור';
  }

  function stopRecording() {
    isRecording = false;
    micBtn.classList.remove('recording');
    micStatus.textContent = 'לחצו כדי לדבר';
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
setupTabs();
setupSettings();
setupConnStatus();
setupRecording();
setupSendDiscard();
setupInstall();
registerServiceWorker();
refreshQueueBadge();
attemptSync();
