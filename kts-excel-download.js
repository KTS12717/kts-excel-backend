/**
 * kts-excel-download.js
 * =====================
 * Drop-in JavaScript module for the KTS Fleet Manager HTML frontend.
 *
 * Provides two public functions:
 *   downloadHA0935(data, monthDate)   — Khan Transportation export
 *   downloadHV0713(data, monthDate)   — Priority Transportation export
 *
 * How it works:
 *   1. Collects data from the app's in-memory state (logs, routes, absences)
 *   2. POSTs JSON to the Python backend
 *   3. Receives a real .xlsx binary stream
 *   4. Triggers a browser file download via a temporary object URL
 *
 * Configuration:
 *   Set BACKEND_URL to your deployed backend address.
 *   During local development use http://localhost:5000
 *   In production use your Render/Railway/Fly.io URL.
 */

// ─── CONFIGURATION ────────────────────────────────────────────────────────────
const BACKEND_URL = window.KTS_BACKEND_URL || "https://your-backend.onrender.com";

// ─── UTILITIES ────────────────────────────────────────────────────────────────

/**
 * Download a Blob as a file.
 * Creates a temporary anchor, clicks it, then removes it.
 */
function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 150);
}

/**
 * POST JSON to a backend endpoint and return the response as a Blob.
 * Throws an Error with the server's error message if the response is not ok.
 */
async function postForBlob(endpoint, payload) {
  const response = await fetch(`${BACKEND_URL}${endpoint}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    let message = `Server error ${response.status}`;
    try {
      const err = await response.json();
      message = err.error || message;
    } catch (_) {}
    throw new Error(message);
  }

  return await response.blob();
}

/**
 * Format a Date as "YYYY-MM-DD" for the backend.
 */
function toISODate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}-01`;
}

// ─── DATA BUILDERS ────────────────────────────────────────────────────────────

/**
 * Build the OVO segment dictionary for a set of routes.
 *
 * @param {Array}  routes       - Array of route objects from app state
 * @param {Array}  logs         - Array of submitted log objects from app state
 * @param {number} year
 * @param {number} month        - 1-indexed
 * @param {Array}  routeNames   - route names to include (e.g. ["Khan1"…] or ["6770","6771"])
 * @returns {Object}  { "Khan1": { "AM": { "2":1, "3":1 … }, "PM": {…} } … }
 */
function buildOvoPayload(routes, logs, year, month, routeNames) {
  const ovo = {};

  routeNames.forEach((routeName) => {
    const route = routes.find(
      (r) => r.name === routeName || r.vehicleId === routeName
    );
    if (!route) return;

    const amDays = {};
    const pmDays = {};

    const daysInMonth = new Date(year, month, 0).getDate();
    for (let day = 1; day <= daysInMonth; day++) {
      const d = new Date(year, month - 1, day);
      if (d.getDay() === 0 || d.getDay() === 6) continue; // skip weekends

      const dateStr = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;

      const amLog = logs.find(
        (l) =>
          (l.routeId === route.id || l.routeName === route.name) &&
          l.date === dateStr &&
          l.period === "AM"
      );
      const pmLog = logs.find(
        (l) =>
          (l.routeId === route.id || l.routeName === route.name) &&
          l.date === dateStr &&
          l.period === "PM"
      );

      amDays[String(day)] = amLog ? 1 : 0;
      pmDays[String(day)] = pmLog ? 1 : 0;
    }

    ovo[routeName] = { AM: amDays, PM: pmDays };
  });

  return ovo;
}

/**
 * Build the SSO miles dictionary.
 *
 * @param {Array}  routes
 * @param {Array}  logs
 * @param {number} year
 * @param {number} month
 * @param {Array}  routeNames
 * @returns {Object}  { "Khan1": { "2": 251.0, "3": 243.0 … } … }
 */
function buildSsoPayload(routes, logs, year, month, routeNames) {
  const sso = {};

  routeNames.forEach((routeName) => {
    const route = routes.find(
      (r) => r.name === routeName || r.vehicleId === routeName
    );
    if (!route) return;

    const dailyMiles = {};
    const daysInMonth = new Date(year, month, 0).getDate();

    for (let day = 1; day <= daysInMonth; day++) {
      const d = new Date(year, month - 1, day);
      if (d.getDay() === 0 || d.getDay() === 6) continue;

      const dateStr = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      const dayLogs = logs.filter(
        (l) =>
          (l.routeId === route.id || l.routeName === route.name) &&
          l.date === dateStr
      );

      if (dayLogs.length > 0) {
        const total = dayLogs.reduce((sum, l) => {
          return sum + (parseFloat(l.totalMiles) || 0);
        }, 0);
        if (total > 0) {
          dailyMiles[String(day)] = Math.round(total * 10) / 10;
        }
      }
    }

    sso[routeName] = dailyMiles;
  });

  return sso;
}

/**
 * Build the attendance grid from logs + absences.
 *
 * @param {Array}  routes
 * @param {Array}  logs
 * @param {Array}  absences    - Array of absence objects from app state
 * @param {number} year
 * @param {number} month
 * @param {number} firstDataRow - first Excel row number (13 for both templates)
 * @returns {Object}  { "13": { "2": 2, "3": 2, "4": "NS" … } … }
 */
function buildAttendancePayload(routes, logs, absences, year, month, firstDataRow = 13) {
  // Collect all participants across all routes, sorted alphabetically by last name
  const allParticipants = [];
  const seen = new Set();

  routes.forEach((route) => {
    (route.clients || []).forEach((client) => {
      if (!seen.has(client.name)) {
        seen.add(client.name);
        const parts = client.name.split(", ");
        allParticipants.push({
          name: client.name,
          lastName: parts[0] || client.name,
          firstName: parts[1] || "",
          routeId: route.id,
          routeName: route.name,
        });
      }
    });
  });

  allParticipants.sort((a, b) => a.lastName.localeCompare(b.lastName));

  const attendance = {};
  const daysInMonth = new Date(year, month, 0).getDate();

  allParticipants.forEach((participant, idx) => {
    const rowNum = firstDataRow + idx;
    const dayCodes = {};

    for (let day = 1; day <= daysInMonth; day++) {
      const d = new Date(year, month - 1, day);
      if (d.getDay() === 0 || d.getDay() === 6) continue;

      const dateStr = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;

      // Check calendar absences first
      const absence = (absences || []).find(
        (a) =>
          a.clientName === participant.name &&
          dateStr >= a.startDate &&
          dateStr <= (a.endDate || a.startDate)
      );
      if (absence) {
        dayCodes[String(day)] = absence.type === "sick" ? "NS" : "C";
        continue;
      }

      // Check log statuses
      const dayLogs = logs.filter(
        (l) =>
          l.date === dateStr &&
          (l.routeId === participant.routeId || l.routeName === participant.routeName)
      );

      if (dayLogs.length === 0) continue;

      const amLog = dayLogs.find((l) => l.period === "AM");
      const pmLog = dayLogs.find((l) => l.period === "PM");
      const getStatus = (log) =>
        log?.clientStatuses?.[participant.name] || null;

      const amStatus = getStatus(amLog);
      const pmStatus = getStatus(pmLog);
      const amPresent = amStatus === "present";
      const pmPresent = pmStatus === "present";
      const isNoShow = amStatus === "noshow" || pmStatus === "noshow";
      const isCancelled = amStatus === "cancelled" || pmStatus === "cancelled";

      if (amPresent && pmPresent) dayCodes[String(day)] = 2;
      else if (amPresent || pmPresent) dayCodes[String(day)] = 1;
      else if (isNoShow) dayCodes[String(day)] = "NS";
      else if (isCancelled) dayCodes[String(day)] = "C";
      else dayCodes[String(day)] = 0;
    }

    if (Object.keys(dayCodes).length > 0) {
      attendance[String(rowNum)] = dayCodes;
    }
  });

  return attendance;
}

// ─── PUBLIC DOWNLOAD FUNCTIONS ────────────────────────────────────────────────

/**
 * Download the HA0935 Khan Transportation Excel file.
 *
 * @param {Object} appState   - The app's current state object containing
 *                              routes, logs, absences, selectedMonth
 * @param {Function} onStart  - Called when download begins (show spinner)
 * @param {Function} onDone   - Called when complete or on error
 *
 * Example usage (in your React component or HTML handler):
 *   downloadHA0935(
 *     { routes, logs, absences, selectedMonth: "2026-03" },
 *     () => setLoading(true),
 *     (err) => { setLoading(false); if (err) alert(err.message); }
 *   );
 */
async function downloadHA0935(appState, onStart, onDone) {
  const { routes, logs, absences, selectedMonth } = appState;
  const [year, month] = selectedMonth.split("-").map(Number);

  const routeNames = [
    "Khan1","Khan2","Khan3","Khan4","Khan5",
    "Khan6","Khan7","Khan8","Khan9","Khan10",
  ];

  const payload = {
    month_date:  toISODate(new Date(year, month - 1, 1)),
    ovo:         buildOvoPayload(routes, logs, year, month, routeNames),
    sso:         buildSsoPayload(routes, logs, year, month, routeNames),
    attendance:  buildAttendancePayload(routes, logs, absences, year, month, 13),
  };

  try {
    if (onStart) onStart();
    const blob = await postForBlob("/api/export/ha0935", payload);
    const monthLabel = new Date(year, month - 1, 1)
      .toLocaleDateString("en-US", { month: "short", year: "numeric" })
      .replace(" ", "");
    triggerDownload(blob, `HA0935_Khan_${monthLabel}.xlsx`);
    if (onDone) onDone(null);
  } catch (err) {
    console.error("HA0935 export failed:", err);
    if (onDone) onDone(err);
  }
}

/**
 * Download the HV0713 Priority Transportation Excel file.
 *
 * @param {Object} appState   - Same structure as above
 * @param {Function} onStart
 * @param {Function} onDone
 */
async function downloadHV0713(appState, onStart, onDone) {
  const { routes, logs, absences, selectedMonth } = appState;
  const [year, month] = selectedMonth.split("-").map(Number);

  // HV0713 uses route vehicleIds "6770" and "6771"
  const priorityRoutes = routes
    .filter((r) => r.vehicleId === "6770" || r.vehicleId === "6771")
    .map((r) => r.vehicleId);

  if (priorityRoutes.length === 0) {
    if (onDone) onDone(new Error("No Priority Transportation routes (6770, 6771) found."));
    return;
  }

  const payload = {
    month_date:  toISODate(new Date(year, month - 1, 1)),
    ovo:         buildOvoPayload(routes, logs, year, month, priorityRoutes),
    sso:         buildSsoPayload(routes, logs, year, month, priorityRoutes),
    attendance:  buildAttendancePayload(routes, logs, absences, year, month, 13),
  };

  try {
    if (onStart) onStart();
    const blob = await postForBlob("/api/export/hv0713", payload);
    const monthLabel = new Date(year, month - 1, 1)
      .toLocaleDateString("en-US", { month: "short", year: "numeric" })
      .replace(" ", "");
    triggerDownload(blob, `HV0713_Priority_${monthLabel}.xlsx`);
    if (onDone) onDone(null);
  } catch (err) {
    console.error("HV0713 export failed:", err);
    if (onDone) onDone(err);
  }
}

// Export for use as a module or global
if (typeof module !== "undefined") {
  module.exports = { downloadHA0935, downloadHV0713 };
}
