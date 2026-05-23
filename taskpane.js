/* global Office */

const STORAGE_KEY = "gemini_api_key";

// ── DOM refs ──────────────────────────────────────────────────────────────────
const apiKeyInput    = document.getElementById("api-key");
const saveKeyBtn     = document.getElementById("save-key-btn");
const keyStatus      = document.getElementById("key-status");
const extractBtn     = document.getElementById("extract-btn");
const loading        = document.getElementById("loading");
const errorBanner    = document.getElementById("error-banner");
const errorMessage   = document.getElementById("error-message");
const resultsSection = document.getElementById("results-section");
const resultsCount   = document.getElementById("results-count");
const eventsList     = document.getElementById("events-list");
const noEvents       = document.getElementById("no-events");

// ── Init ──────────────────────────────────────────────────────────────────────
Office.onReady(() => {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    apiKeyInput.value = saved;
    setKeyStatus("API key loaded", false);
  }

  saveKeyBtn.addEventListener("click", saveApiKey);
  extractBtn.addEventListener("click", handleExtract);
});

// ── API key ───────────────────────────────────────────────────────────────────
function saveApiKey() {
  const key = apiKeyInput.value.trim();
  if (!key) {
    setKeyStatus("Please enter an API key", true);
    return;
  }
  localStorage.setItem(STORAGE_KEY, key);
  setKeyStatus("Saved!", false);
}

function setKeyStatus(msg, isError) {
  keyStatus.textContent = msg;
  keyStatus.className = "status-text" + (isError ? " error" : "");
}

// ── Main flow ─────────────────────────────────────────────────────────────────
async function handleExtract() {
  hideAll();

  const apiKey = localStorage.getItem(STORAGE_KEY) || apiKeyInput.value.trim();
  if (!apiKey) {
    showError("Please enter and save your Gemini API key first.");
    return;
  }

  showLoading(true);
  extractBtn.disabled = true;

  try {
    const emailText = await getEmailText();
    const events    = await extractEvents(apiKey, emailText);
    renderEvents(events);
  } catch (err) {
    showError(err.message || "An unexpected error occurred.");
  } finally {
    showLoading(false);
    extractBtn.disabled = false;
  }
}

// ── Read email ────────────────────────────────────────────────────────────────
function getEmailText() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error("No email item found. Open an email and try again."));
      return;
    }

    // Collect subject + sender for context
    const subject = item.subject || "";
    const sender  = item.from ? `${item.from.displayName} <${item.from.emailAddress}>` : "";

    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error("Could not read email body: " + result.error.message));
        return;
      }
      const body = result.value || "";
      resolve(`Subject: ${subject}\nFrom: ${sender}\n\n${body}`);
    });
  });
}

// ── Call Gemini ───────────────────────────────────────────────────────────────
async function extractEvents(apiKey, emailText) {
  const systemPrompt = `You are a calendar assistant. Given the text of an email, extract all calendar events mentioned.

For each event return a JSON object with these fields:
- title: short event title (required)
- date: ISO 8601 date string if determinable, otherwise a human-readable string like "Next Monday" (required)
- startTime: time in HH:MM 24h format, or null if unknown
- endTime: time in HH:MM 24h format, or null if unknown
- location: physical or virtual location, or null
- description: 1-3 sentence summary of what the event is about
- attendees: array of names/emails if mentioned, otherwise []

Return a JSON array of events. If there are no events, return an empty array [].
Do not include any text outside the JSON array.`;

  const model = "gemini-2.5-flash";
  const url   = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      system_instruction: {
        parts: [{ text: systemPrompt }],
      },
      contents: [
        { role: "user", parts: [{ text: emailText }] },
      ],
      generationConfig: {
        temperature: 0,
        responseMimeType: "application/json",
      },
    }),
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    const msg = err.error?.message || `Gemini API error (${response.status})`;
    throw new Error(msg);
  }

  const data    = await response.json();
  const content = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || "[]";

  try {
    const parsed = JSON.parse(content);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    throw new Error("Could not parse event data from Gemini response.");
  }
}

// ── Render ────────────────────────────────────────────────────────────────────
function renderEvents(events) {
  if (!events.length) {
    noEvents.classList.remove("hidden");
    return;
  }

  resultsCount.textContent =
    events.length === 1 ? "1 Event Found" : `${events.length} Events Found`;
  resultsSection.classList.remove("hidden");
  eventsList.innerHTML = "";

  events.forEach((ev, idx) => {
    eventsList.appendChild(buildEventCard(ev, idx));
  });
}

function buildEventCard(ev, idx) {
  const card = document.createElement("div");
  card.className = "event-card";
  card.dataset.idx = idx;

  const timeStr = formatTimeRange(ev.startTime, ev.endTime);

  card.innerHTML = `
    <div class="event-title">${escHtml(ev.title)}</div>
    <div class="event-meta">
      <div class="event-meta-row">
        <span class="label">Date</span>
        <span>${escHtml(ev.date || "Unknown")}</span>
      </div>
      ${timeStr ? `
      <div class="event-meta-row">
        <span class="label">Time</span>
        <span>${escHtml(timeStr)}</span>
      </div>` : ""}
      ${ev.location ? `
      <div class="event-meta-row">
        <span class="label">Location</span>
        <span>${escHtml(ev.location)}</span>
      </div>` : ""}
      ${ev.attendees?.length ? `
      <div class="event-meta-row">
        <span class="label">Attendees</span>
        <span>${escHtml(ev.attendees.join(", "))}</span>
      </div>` : ""}
    </div>
    ${ev.description ? `<div class="event-description">${escHtml(ev.description)}</div>` : ""}
    <div class="event-actions">
      <button class="add-btn" data-idx="${idx}">Add to Calendar</button>
    </div>
  `;

  card.querySelector(".add-btn").addEventListener("click", () => addToCalendar(ev, card));
  return card;
}

function formatTimeRange(start, end) {
  if (!start) return null;
  const s = formatTime(start);
  const e = end ? formatTime(end) : null;
  return e ? `${s} – ${e}` : s;
}

function formatTime(t) {
  const [h, m] = t.split(":").map(Number);
  const ampm = h >= 12 ? "PM" : "AM";
  const hour  = h % 12 || 12;
  return `${hour}:${String(m).padStart(2, "0")} ${ampm}`;
}

// ── Add to calendar ───────────────────────────────────────────────────────────
function addToCalendar(ev, card) {
  const isAllDay = !ev.startTime;

  let startDt, endDt;
  if (isAllDay) {
    // All-day: midnight → midnight next day
    startDt = buildDateOnly(ev.date);
    endDt   = new Date(startDt);
    endDt.setDate(endDt.getDate() + 1);
  } else {
    startDt = buildDateTime(ev.date, ev.startTime);
    endDt   = ev.endTime ? buildDateTime(ev.date, ev.endTime) : addOneHour(startDt);
  }

  const appointmentDetails = {
    subject:  ev.title,
    body:     ev.description || "",
    start:    startDt,
    end:      endDt,
    location: ev.location || "",
    requiredAttendees: [],
  };

  // Try Office JS newAppointment first (requires Mailbox 1.1)
  if (Office.context.mailbox.displayNewAppointmentForm) {
    Office.context.mailbox.displayNewAppointmentForm(appointmentDetails);
    markAdded(card);
  } else {
    // Fallback: open Outlook web new event URL
    const url = buildOutlookWebUrl(ev, startDt, endDt, isAllDay);
    window.open(url, "_blank");
    markAdded(card);
  }
}

/** Parse date string to midnight local time (for all-day events). */
function buildDateOnly(dateStr) {
  const base = new Date(dateStr);
  if (isNaN(base.getTime())) {
    // Relative/unrecognised date — use today
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return today;
  }
  base.setHours(0, 0, 0, 0);
  return base;
}

function buildDateTime(dateStr, timeStr) {
  if (!dateStr) return new Date();
  const base = new Date(dateStr);
  if (isNaN(base.getTime())) {
    const fallback = new Date();
    if (timeStr) applyTime(fallback, timeStr);
    return fallback;
  }
  if (timeStr) applyTime(base, timeStr);
  return base;
}

function applyTime(dt, timeStr) {
  const [h, m] = timeStr.split(":").map(Number);
  dt.setHours(h, m, 0, 0);
}

function addOneHour(dt) {
  const d = new Date(dt);
  d.setHours(d.getHours() + 1);
  return d;
}

function buildOutlookWebUrl(ev, startDt, endDt, isAllDay) {
  const params = new URLSearchParams({
    subject:  ev.title    || "",
    body:     ev.description || "",
    location: ev.location || "",
  });
  if (isAllDay) {
    // Outlook Web uses date-only strings (YYYY-MM-DD) for all-day events
    params.set("startdt", formatDateOnly(startDt));
    params.set("enddt",   formatDateOnly(endDt));
    params.set("allday",  "true");
  } else {
    params.set("startdt", startDt.toISOString());
    params.set("enddt",   endDt.toISOString());
  }
  return `https://outlook.live.com/calendar/0/action/compose?${params.toString()}`;
}

function formatDateOnly(dt) {
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, "0");
  const d = String(dt.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function markAdded(card) {
  card.classList.add("added");
  const actionsDiv = card.querySelector(".event-actions");
  actionsDiv.innerHTML = `<div class="added-label">Added to Calendar</div>`;
}

// ── UI helpers ────────────────────────────────────────────────────────────────
function hideAll() {
  loading.classList.add("hidden");
  errorBanner.classList.add("hidden");
  resultsSection.classList.add("hidden");
  noEvents.classList.add("hidden");
}

function showLoading(show) {
  loading.classList.toggle("hidden", !show);
}

function showError(msg) {
  errorMessage.textContent = msg;
  errorBanner.classList.remove("hidden");
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
