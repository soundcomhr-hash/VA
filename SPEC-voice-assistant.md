# SPEC — "העוזרת" Voice Assistant PWA (v1)

## Who this is for
Owner: Ziv Weisberg, music-therapy business, Israel. Single user in v1 (Ziv only).
Android phone. Hebrew speaker. Not a developer — all explanations must be simple.
Time is the scarce resource; tokens/compute are cheap.

## Hard constraints (non-negotiable)
1. **₪0/month.** No paid APIs, no paid hosting, no subscriptions.
2. **No dependency on external AI services in the critical path.** No Gemini/OpenAI/
   Claude API calls for core function. (Google infrastructure like Apps Script,
   Sheets, Calendar, GitHub Pages is acceptable — the ban is on AI-model APIs whose
   policy may change.)
3. **Everything owned by Ziv:** his GitHub repo, his Google account, his data.
4. **Nothing is written to the calendar or any ledger without explicit confirmation**
   (tap אשר) — except pure capture (raw recordings log), which is always saved.
5. Data must remain human-readable in Google Sheets tabs (the "legacy ledger" —
   if all code dies, the sheets still tell the truth).

## Architecture
- **Frontend:** static PWA, hosted on GitHub Pages (Ziv's account). Installable on
  Android (manifest + service worker), full offline support with queued sync.
  RTL, Hebrew UI.
- **Backend:** Google Apps Script web app (extends Ziv's existing "SoundCom Ops"
  project — it already has tabs and a session engine). JSON over POST. Shared
  secret token stored in the app (localStorage) and in Script Properties.
- **Transcription:** Web Speech API (on-device Hebrew via Gboard/Android) for live
  speech-to-text INSIDE the app. Fallback: MediaRecorder captures audio; if speech
  recognition unavailable/failed, the audio is stored (Drive via backend) and the
  entry is marked "לתמלול" in the inbox. ANSWERED by Ziv's device tests:
  (a) live Hebrew voice typing WORKS but quality is mediocre → audio is ALWAYS
  recorded and kept as source of truth; transcript is an editable draft; bad
  transcripts go to inbox as "לתמלול" with audio attached. (b) YES, true
  no-signal situations exist → offline recording + queued sync is REQUIRED in
  milestone 1. Future option (not v1): self-hosted ivrit.ai Whisper to
  re-transcribe the audio queue.
- **Parser (the brain, v1):** deterministic Hebrew slot-filler running in Apps
  Script. Closed vocabulary (below). No ML. When parsing fails or slots are
  missing → item goes to inbox with a generated follow-up question. The parser
  must NEVER guess money, dates, or people; missing = ask.

## Screens (PWA, bottom tabs)
1. **הקלטה** — one huge mic button. Tap to talk; live transcript appears in an
   editable text box; big "שלח" button. Default flow = speak → auto-send after
   confirm tap; text is editable before send if wanted ("both" mode).
2. **ממתין** — the inbox. Cards: parsed interpretation in Hebrew + buttons:
   [אשר] [ערוך] [מחק], or the follow-up question with a quick answer field.
   Confirmed calendar actions call the backend which writes to Google Calendar.
3. **היום** — today's calendar events (read from backend) with status colors:
   green = confirmed/handled, orange = pending, red = passed with no confirmation.
   (v1 = today only; month grid view is v2.)

## Backend endpoints (Apps Script doPost, action field)
- `capture` {text, audioBase64?} → logs raw entry to tab "קלט", runs parser,
  creates inbox item(s) in tab "ממתין". Returns parse result.
- `inbox_list` / `inbox_answer` {id, answer} / `inbox_confirm` {id} /
  `inbox_delete` {id}
- `today` → merged view: Google Calendar events for today (CalendarApp) +
  confirmation status from "מפגשים" tab.
- On `inbox_confirm` of a calendar action → CalendarApp create/update event.
- Session-attendance confirmations continue to use the existing confirmation
  page/tab (do not rebuild it here; just link to it from היום).

## Parser vocabulary (seed — keep in a Sheet tab "מילון" so Ziv can extend without code)
- **People:** זיו, חן (גבע), ספיר, נווה (new teacher), אבי.
- **Places:** בית מיכל (ראשון לציון), מעון רוחמה (כפר סבא), ארקדש (יהוד),
  נווה האירוס, אקים-כלנית, מעון הוד, הגנים, בית קסלר, טירת הכרמל,
  תיכון חדש (inactive), נהריה (closed).
- **Intents:** קבע/תקבע (schedule), תזכיר/תזכורת (reminder), העבר/הזז (move event),
  בטל (cancel), היה/היו/נכחו (attendance), נתתי/קיבל ציוד (equipment transfer),
  קניתי (purchase→inventory), משימה/צריך (task).
- **Time expressions:** Hebrew days (ראשון..שבת), "מחר", "מחרתיים", "שבוע הבא",
  hours like "10 וחצי", "3 וחצי", "רבע ל...". Resolve relative to Asia/Jerusalem.
- Fuzzy place matching required (voice typos: "באר קדש"→ארקדש, "דוי האירוס"→נווה האירוס).

## Behavior rules (from Ziv, verbatim intent)
- Complete info (day+time+place) → prepare action → inbox → one tap אשר → calendar.
- Incomplete → NEVER write, NEVER drop: inbox item + one precise follow-up question.
- Equipment given to a person → pending item assigned to that person's confirmation
  (v1: Ziv confirms on their behalf; multi-user is v2), then inventory deduction
  in tab "מלאי" + auto reminder event "לבדוק ציוד אצל X" (+7 days default).
- Every raw recording is preserved in "קלט" with timestamp (audit trail; OCD-trust).

## Existing assets (context)
- Google Calendar (primary) already has recurring 🎵 session events (Mon בית מיכל
  15:30, Tue אקים-כלנית 10:00, Wed רוחמה 15:30, Thu ארקדש 10:30, Thu נווה האירוס
  13:00) and task-block events. 🎵 prefix marks session events.
- Apps Script project "SoundCom Ops" with tabs: עובדים, לקוחות, מקומות, מטופלים,
  לוז, מפגשים, מלאי, הגדרות + a session-confirmation web page (patient chips).
- Airtable payment base and Morning (Green Invoice) integration exist but are
  OUT OF SCOPE here. Do not touch.

## Milestones (build in this order, each independently testable)
1. PWA shell on GitHub Pages: record button, live Hebrew transcription, editable
   text, offline queue, send to a stub endpoint. Installable on Ziv's phone.
2. Backend capture + "קלט" log + inbox CRUD + ממתין screen.
3. Parser v1 (vocabulary tab, intents, time resolution, fuzzy places) +
   follow-up-question generation.
4. Calendar actions on confirm (create/move/cancel event, reminders).
5. היום screen with status colors; equipment/inventory flow.

## Working agreement for Claude Code
- Explain every step in plain language; Ziv runs commands you give him verbatim.
- Small commits per milestone; each milestone ends with "how to test on your phone".
- Any design ambiguity: ASK, don't assume — especially anything touching money,
  attendance, or calendar writes.
