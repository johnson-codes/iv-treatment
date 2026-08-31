/**
 * IV Treatment intake form.
 * Paste the Google Apps Script Web App URL below after you deploy the backend.
 */
const WEBHOOK_URL = "";

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const PACKAGES = [
  "Hydration Boost",
  "Immunity Glow",
  "NAD+ Infusion",
  "Energy & Performance",
];
const TIME_WINDOWS = ["Morning", "Afternoon", "Evening"];

const form = document.getElementById("intake-form");
const submitBtn = document.getElementById("submit-btn");
const formError = document.getElementById("form-error");
const formPanel = document.getElementById("form-panel");
const successPanel = document.getElementById("success-panel");
const submissionIdEl = document.getElementById("submission-id");
const dateInput = document.getElementById("preferredDate");

let submitting = false;

function todayISO() {
  const now = new Date();
  const offset = now.getTimezoneOffset();
  const local = new Date(now.getTime() - offset * 60 * 1000);
  return local.toISOString().slice(0, 10);
}

dateInput.min = todayISO();
dateInput.value = dateInput.value || todayISO();

function getUtm() {
  const params = new URLSearchParams(window.location.search);
  return {
    utm_source: (params.get("utm_source") || "").trim(),
    utm_medium: (params.get("utm_medium") || "").trim(),
    utm_campaign: (params.get("utm_campaign") || "").trim(),
  };
}

function digitsOnly(value) {
  return (value || "").replace(/\D/g, "");
}

function setError(name, message) {
  const field =
    form.querySelector(`[name="${name}"]`) || document.getElementById(name);
  const errorEl = form.querySelector(`[data-error-for="${name}"]`);
  if (errorEl) {
    errorEl.textContent = message || "";
    errorEl.classList.toggle("hidden", !message);
  }
  if (field) {
    if (field.type === "radio" || field.type === "checkbox") {
      const group = form.querySelectorAll(`[name="${name}"]`);
      group.forEach((el) => el.setAttribute("aria-invalid", message ? "true" : "false"));
    } else {
      field.classList.toggle("border-red-400", Boolean(message));
      field.setAttribute("aria-invalid", message ? "true" : "false");
    }
  }
}

function clearErrors() {
  form.querySelectorAll("[data-error-for]").forEach((el) => {
    el.textContent = "";
    el.classList.add("hidden");
  });
  form.querySelectorAll("[aria-invalid]").forEach((el) => {
    el.setAttribute("aria-invalid", "false");
    el.classList.remove("border-red-400");
  });
  formError.classList.add("hidden");
  formError.textContent = "";
}

function readForm() {
  const data = new FormData(form);
  return {
    fullName: (data.get("fullName") || "").toString().trim(),
    phone: (data.get("phone") || "").toString().trim(),
    email: (data.get("email") || "").toString().trim(),
    package: (data.get("package") || "").toString(),
    preferredDate: (data.get("preferredDate") || "").toString(),
    timeWindow: (data.get("timeWindow") || "").toString(),
    healthNotes: (data.get("healthNotes") || "").toString().trim(),
    consent: form.querySelector("#consent").checked,
    honeypot: (data.get("company_website") || "").toString().trim(),
  };
}

function validate(values) {
  const errors = {};

  if (values.fullName.length < 2) {
    errors.fullName = "Please enter your full name.";
  }
  const phoneDigits = digitsOnly(values.phone);
  if (phoneDigits.length < 10 || phoneDigits.length > 15) {
    errors.phone = "Enter a valid phone number.";
  }
  if (!EMAIL_RE.test(values.email)) {
    errors.email = "Enter a valid email address.";
  }
  if (!PACKAGES.includes(values.package)) {
    errors.package = "Select a package.";
  }
  if (!values.preferredDate) {
    errors.preferredDate = "Choose a preferred date.";
  } else if (values.preferredDate < todayISO()) {
    errors.preferredDate = "Please choose today or a future date.";
  }
  if (!TIME_WINDOWS.includes(values.timeWindow)) {
    errors.timeWindow = "Select a time window.";
  }
  if (!values.consent) {
    errors.consent = "Consent is required so we can confirm your booking.";
  }

  return errors;
}

function showErrors(errors) {
  Object.entries(errors).forEach(([name, message]) => setError(name, message));
  const first = Object.keys(errors)[0];
  if (!first) return;
  const target =
    form.querySelector(`[name="${first}"]`) || document.getElementById(first);
  if (target && typeof target.focus === "function") target.focus();
}

function setLoading(isLoading) {
  submitting = isLoading;
  submitBtn.disabled = isLoading;
  submitBtn.setAttribute("aria-busy", isLoading ? "true" : "false");
  submitBtn.textContent = isLoading ? "Sending…" : "Request my session";
}

function showSuccess(id) {
  formPanel.classList.add("hidden");
  successPanel.classList.remove("hidden");
  submissionIdEl.textContent = id;
}

function resetToForm() {
  form.reset();
  dateInput.min = todayISO();
  dateInput.value = todayISO();
  clearErrors();
  successPanel.classList.add("hidden");
  formPanel.classList.remove("hidden");
}

function buildPayload(values) {
  const utm = getUtm();
  return {
    fullName: values.fullName,
    phone: values.phone,
    email: values.email.toLowerCase(),
    package: values.package,
    preferredDate: values.preferredDate,
    timeWindow: values.timeWindow,
    healthNotes: values.healthNotes,
    consent: true,
    ...utm,
    pageUrl: window.location.href.split("#")[0],
  };
}

async function sendLead(payload) {
  if (!WEBHOOK_URL) {
    await new Promise((resolve) => setTimeout(resolve, 700));
    return { ok: true, id: "IV-DEMO-0001", demo: true };
  }

  const response = await fetch(WEBHOOK_URL, {
    method: "POST",
    redirect: "follow",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    throw new Error("Request failed");
  }

  const data = await response.json();
  if (!data || data.ok !== true) {
    throw new Error(data && data.error ? data.error : "Request failed");
  }
  return data;
}

form.addEventListener("blur", (event) => {
  const name = event.target && event.target.name;
  if (!name || name === "company_website" || name === "healthNotes") return;
  const values = readForm();
  const errors = validate(values);
  setError(name, errors[name] || "");
}, true);

form.addEventListener("change", (event) => {
  const name = event.target && event.target.name;
  if (!name) return;
  const values = readForm();
  const errors = validate(values);
  if (!errors[name]) setError(name, "");
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (submitting) return;

  clearErrors();
  const values = readForm();

  if (values.honeypot) {
    showSuccess("IV-OK");
    return;
  }

  const errors = validate(values);
  if (Object.keys(errors).length) {
    showErrors(errors);
    return;
  }

  setLoading(true);
  try {
    const result = await sendLead(buildPayload(values));
    showSuccess(result.id || "received");
    form.reset();
    dateInput.value = todayISO();
  } catch (err) {
    formError.textContent =
      "We couldn’t send your request. Please try again in a moment.";
    formError.classList.remove("hidden");
  } finally {
    setLoading(false);
  }
});

document.getElementById("another-request").addEventListener("click", resetToForm);
