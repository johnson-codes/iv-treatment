/**
 * Smile Well IV membership intake.
 * Paste the Google Apps Script Web App URL below after you deploy the backend.
 */
const WEBHOOK_URL = "";

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const MIN_AGE = 18;

const form = document.getElementById("intake-form");
const submitBtn = document.getElementById("submit-btn");
const formError = document.getElementById("form-error");
const formPanel = document.getElementById("form-panel");
const successPanel = document.getElementById("success-panel");
const submissionIdEl = document.getElementById("submission-id");
const successTermEl = document.getElementById("success-term");
const dobInput = document.getElementById("dob");
const startInput = document.getElementById("startDate");
const expiryInput = document.getElementById("expiryDate");

let submitting = false;
let lastTerm = "";

function todayISO() {
  const now = new Date();
  const offset = now.getTimezoneOffset();
  const local = new Date(now.getTime() - offset * 60 * 1000);
  return local.toISOString().slice(0, 10);
}

function pad(n) {
  return String(n).padStart(2, "0");
}

function addYearsISO(iso, years) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-").map(Number);
  const end = new Date(y + years, m - 1, d);
  if (end.getMonth() !== m - 1) end.setDate(0);
  return `${end.getFullYear()}-${pad(end.getMonth() + 1)}-${pad(end.getDate())}`;
}

function ageOn(dobISO, onISO) {
  const [y, m, d] = dobISO.split("-").map(Number);
  const [oy, om, od] = onISO.split("-").map(Number);
  let age = oy - y;
  if (om < m || (om === m && od < d)) age -= 1;
  return age;
}

function namesMatch(a, b) {
  const norm = (s) => s.toLowerCase().replace(/\s+/g, " ").trim();
  return norm(a) === norm(b);
}

function formatDisplayDate(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, m - 1, d).toLocaleDateString("en-CA", {
    year: "numeric",
    month: "short",
    day: "numeric",
  });
}

function syncExpiry() {
  expiryInput.value = addYearsISO(startInput.value, 1);
}

const today = todayISO();
dobInput.max = today;
startInput.min = today;
startInput.value = startInput.value || today;
syncExpiry();
startInput.addEventListener("change", syncExpiry);
startInput.addEventListener("input", syncExpiry);

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
    if (field.type === "checkbox") {
      field.setAttribute("aria-invalid", message ? "true" : "false");
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
    dob: (data.get("dob") || "").toString(),
    phone: (data.get("phone") || "").toString().trim(),
    email: (data.get("email") || "").toString().trim(),
    startDate: (data.get("startDate") || "").toString(),
    expiryDate: addYearsISO((data.get("startDate") || "").toString(), 1),
    healthNotes: (data.get("healthNotes") || "").toString().trim(),
    signature: (data.get("signature") || "").toString().trim(),
    agreement: form.querySelector("#agreement").checked,
    cardAck: form.querySelector("#cardAck").checked,
    consent: form.querySelector("#consent").checked,
    honeypot: (data.get("company_website") || "").toString().trim(),
  };
}

function validate(values) {
  const errors = {};

  if (values.fullName.length < 2) {
    errors.fullName = "Please enter your full name.";
  }
  if (!values.dob) {
    errors.dob = "Enter your date of birth.";
  } else if (values.dob > todayISO()) {
    errors.dob = "Date of birth cannot be in the future.";
  } else if (ageOn(values.dob, todayISO()) < MIN_AGE) {
    errors.dob = "Members must be 18 or older.";
  }
  const phoneDigits = digitsOnly(values.phone);
  if (phoneDigits.length < 10 || phoneDigits.length > 15) {
    errors.phone = "Enter a valid phone number.";
  }
  if (!EMAIL_RE.test(values.email)) {
    errors.email = "Enter a valid email address.";
  }
  if (!values.startDate) {
    errors.startDate = "Choose a start date.";
  } else if (values.startDate < todayISO()) {
    errors.startDate = "Start date must be today or later.";
  }
  if (!values.signature) {
    errors.signature = "Type your full name as your signature.";
  } else if (!namesMatch(values.signature, values.fullName)) {
    errors.signature = "Signature must match your full name.";
  }
  if (!values.agreement) {
    errors.agreement = "Please agree to the membership terms.";
  }
  if (!values.cardAck) {
    errors.cardAck = "Please acknowledge that a card on file is required.";
  }
  if (!values.consent) {
    errors.consent = "Consent is required so we can confirm your membership.";
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
  submitBtn.textContent = isLoading ? "Sending…" : "Submit membership agreement";
}

function showSuccess(id, values) {
  formPanel.classList.add("hidden");
  successPanel.classList.remove("hidden");
  submissionIdEl.textContent = id;
  lastTerm = `${formatDisplayDate(values.startDate)} – ${formatDisplayDate(values.expiryDate)}`;
  successTermEl.textContent = lastTerm;
}

function resetToForm() {
  form.reset();
  startInput.min = todayISO();
  startInput.value = todayISO();
  dobInput.max = todayISO();
  syncExpiry();
  clearErrors();
  successPanel.classList.add("hidden");
  formPanel.classList.remove("hidden");
}

function buildPayload(values) {
  const utm = getUtm();
  return {
    fullName: values.fullName,
    dob: values.dob,
    phone: values.phone,
    email: values.email.toLowerCase(),
    startDate: values.startDate,
    expiryDate: values.expiryDate,
    healthNotes: values.healthNotes,
    signature: values.signature,
    agreement: true,
    cardAck: true,
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
  if (!name || name === "company_website" || name === "healthNotes" || name === "expiryDate") return;
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
    showSuccess("IV-OK", values);
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
    showSuccess(result.id || "received", values);
    form.reset();
    startInput.value = todayISO();
    syncExpiry();
  } catch (err) {
    formError.textContent =
      "We couldn’t send your agreement. Please try again in a moment.";
    formError.classList.remove("hidden");
  } finally {
    setLoading(false);
  }
});

document.getElementById("another-request").addEventListener("click", resetToForm);
