const DEFAULT_DL = new URLSearchParams(window.location.search).get("dl") || "";
const OUTLOOK_COMPOSE_BASE = "https://outlook.office.com/calendar/deeplink/compose";

const $ = (id) => document.getElementById(id);

function setDefaultDates() {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, "0");
  const dd = String(today.getDate()).padStart(2, "0");
  const todayString = `${yyyy}-${mm}-${dd}`;

  $("startDate").value = todayString;
  $("endDate").value = todayString;
  $("partialDate").value = todayString;
  $("startTime").value = "09:00";
  $("endTime").value = "17:00";
  $("dl").value = DEFAULT_DL;
}

function formatDateForBody(dateString) {
  if (!dateString) return "";
  const [year, month, day] = dateString.split("-").map(Number);
  return new Date(year, month - 1, day).toLocaleDateString(undefined, {
    weekday: "long",
    year: "numeric",
    month: "long",
    day: "numeric"
  });
}

function formatTimeForBody(timeString) {
  if (!timeString) return "";
  const [hour, minute] = timeString.split(":").map(Number);
  return new Date(2000, 0, 1, hour, minute).toLocaleTimeString(undefined, {
    hour: "numeric",
    minute: "2-digit"
  });
}

function addDays(dateString, days) {
  const [year, month, day] = dateString.split("-").map(Number);
  const date = new Date(year, month - 1, day);
  date.setDate(date.getDate() + days);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function buildOutlookUrl(params) {
  const query = new URLSearchParams();

  Object.entries(params).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== "") {
      query.set(key, value);
    }
  });

  return `${OUTLOOK_COMPOSE_BASE}?${query.toString()}`;
}

function buildLinks() {
  const name = $("name").value.trim() || "Your Name";
  const dl = $("dl").value.trim();
  const allDay = $("allDay").checked;
  const subject = `${name} - OOO`;

  let startdt;
  let enddt;
  let bodyDetails;
  let allDayValue;

  if (allDay) {
    const startDate = $("startDate").value;
    const endDate = $("endDate").value;
    if (!startDate || !endDate) {
      alert("Please choose a start and end date.");
      return;
    }

    startdt = startDate;
    enddt = addDays(endDate, 1);
    allDayValue = "true";

    if (startDate === endDate) {
      bodyDetails = `${name} will be out of office on ${formatDateForBody(startDate)}.`;
    } else {
      bodyDetails = `${name} will be out of office from ${formatDateForBody(startDate)} through ${formatDateForBody(endDate)}.`;
    }
  } else {
    const date = $("partialDate").value;
    const startTime = $("startTime").value;
    const endTime = $("endTime").value;
    const tz = $("timeZone").value.trim();

    if (!date || !startTime || !endTime) {
      alert("Please choose a date, start time, and end time.");
      return;
    }

    startdt = `${date}T${startTime}:00`;
    enddt = `${date}T${endTime}:00`;
    allDayValue = "false";
    bodyDetails = `${name} will be out of office on ${formatDateForBody(date)} from ${formatTimeForBody(startTime)} to ${formatTimeForBody(endTime)}${tz ? ` ${tz}` : ""}.`;
  }

  const body = "";

  const sharedBody = `(Invite ${dl} and uncheck "Request Responses". Delete this before sending.)`

  const sharedUrl = buildOutlookUrl({
    subject,
    body: sharedBody,
    startdt,
    enddt,
    allday: allDayValue,
    freebusy: "free"
  });

  const personalUrl = buildOutlookUrl({
    subject,
    body,
    startdt,
    enddt,
    allday: allDayValue,
    freebusy: "busy"
  });

  $("sharedLink").href = sharedUrl;
  $("sharedUrl").textContent = sharedUrl;
  $("personalLink").href = personalUrl;
  $("personalUrl").textContent = personalUrl;
  $("results").hidden = false;
}

function copyText(elementId) {
  const text = $(elementId).textContent;
  navigator.clipboard.writeText(text).catch(() => {
    alert("Could not copy automatically. You can manually copy the link text below the button.");
  });
}

$("allDay").addEventListener("change", () => {
  const allDay = $("allDay").checked;
  $("allDayFields").hidden = !allDay;
  $("partialDayFields").hidden = allDay;
});

$("buildLinks").addEventListener("click", buildLinks);

$("resetForm").addEventListener("click", () => {
  $("name").value = "";
  $("results").hidden = true;
  $("allDay").checked = true;
  $("allDayFields").hidden = false;
  $("partialDayFields").hidden = true;
  setDefaultDates();
});

$("copyShared").addEventListener("click", () => copyText("sharedUrl"));
$("copyPersonal").addEventListener("click", () => copyText("personalUrl"));

setDefaultDates();
