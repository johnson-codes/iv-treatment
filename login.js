/**
 * Smile Well staff sign-in. After a successful login, open the reception desk.
 */
const firebaseConfig = {
  apiKey: "AIzaSyDgV0ZN5h1MWzQWNwNqe-ZJmy2aBWL8diI",
  authDomain: "smile-well-34579.firebaseapp.com",
  projectId: "smile-well-34579",
  storageBucket: "smile-well-34579.firebasestorage.app",
  messagingSenderId: "462381026785",
  appId: "1:462381026785:web:7e35027eb715c5199616a0",
  measurementId: "G-B67NQ5PX1M",
};

firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();

const authLoading = document.getElementById("auth-loading");
const loginApp = document.getElementById("login-app");
const loginForm = document.getElementById("login-form");
const loginEmail = document.getElementById("login-email");
const loginPassword = document.getElementById("login-password");
const loginError = document.getElementById("login-error");
const loginSubmit = document.getElementById("login-submit");

let signingIn = false;

function authErrorMessage(err) {
  const code = err && err.code ? String(err.code) : "";
  if (code === "auth/invalid-email") return "Enter a valid email address.";
  if (code === "auth/invalid-credential" || code === "auth/wrong-password" || code === "auth/user-not-found") {
    return "Email or password is incorrect.";
  }
  if (code === "auth/too-many-requests") return "Too many attempts. Try again in a few minutes.";
  if (code === "auth/network-request-failed") return "Network error. Check your connection and try again.";
  return err && err.message ? err.message : "Could not sign in.";
}

function showLoginError(message) {
  if (!loginError) return;
  if (!message) {
    loginError.textContent = "";
    loginError.classList.add("hidden");
    return;
  }
  loginError.textContent = message;
  loginError.classList.remove("hidden");
}

function setSigningIn(busy) {
  signingIn = busy;
  if (loginSubmit) {
    loginSubmit.disabled = busy;
    loginSubmit.textContent = busy ? "Signing in…" : "Sign in";
  }
}

function showLoginForm() {
  if (authLoading) authLoading.classList.add("hidden");
  if (loginApp) loginApp.classList.add("is-ready");
}

function goToDesk() {
  location.replace("receptionist.html");
}

auth.onAuthStateChanged((user) => {
  if (user) {
    goToDesk();
    return;
  }
  showLoginForm();
});

if (loginForm) {
  loginForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (signingIn) return;
    const email = loginEmail ? loginEmail.value.trim() : "";
    const password = loginPassword ? loginPassword.value : "";
    if (!email || !password) {
      showLoginError("Enter your email and password.");
      return;
    }
    showLoginError("");
    setSigningIn(true);
    try {
      await auth.signInWithEmailAndPassword(email, password);
      goToDesk();
    } catch (err) {
      showLoginError(authErrorMessage(err));
      setSigningIn(false);
    }
  });
}
