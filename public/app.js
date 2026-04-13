const messagesEl = document.getElementById("messages");
const inputEl = document.getElementById("messageInput");
const sendBtn = document.getElementById("sendBtn");
const coachStartBtn = document.getElementById("coachStartBtn");
const micBtn = document.getElementById("micBtn");
const stopSpeakBtn = document.getElementById("stopSpeakBtn");
const resetBtn = document.getElementById("resetBtn");
const autoSpeakEl = document.getElementById("autoSpeak");
const autoSendVoiceEl = document.getElementById("autoSendVoice");
const statusEl = document.getElementById("status");
const coachVisualEl = document.getElementById("coachVisual");
const coachVisualStatusEl = document.getElementById("coachVisualStatus");
const coachAvatarEl = document.getElementById("coachAvatar");
const coachMouthEl = document.getElementById("coachMouth");
const speechVoiceEl = document.getElementById("speechVoice");

const mainHeader = document.getElementById("mainHeader");
const splashPage = document.getElementById("splashPage");
const getStartedBtn = document.getElementById("getStartedBtn");
const authPanel = document.getElementById("authPanel");
const accountPanel = document.getElementById("accountPanel");
const dashboardPage = document.getElementById("dashboardPage");
const chatPage = document.getElementById("chatPage");
const authStatusEl = document.getElementById("authStatus");
const profileStatusEl = document.getElementById("profileStatus");
const passwordStatusEl = document.getElementById("passwordStatus");
const showLoginBtn = document.getElementById("showLoginBtn");
const showRegisterBtn = document.getElementById("showRegisterBtn");
const loginForm = document.getElementById("loginForm");
const registerForm = document.getElementById("registerForm");
const logoutBtn = document.getElementById("logoutBtn");
const passwordRequiredBanner = document.getElementById("passwordRequiredBanner");
const accountMenuBtn = document.getElementById("accountMenuBtn");
const adminMenuBtn = document.getElementById("adminMenuBtn");
const dashboardMenuBtn = document.getElementById("dashboardMenuBtn");
const backToChatBtn = document.getElementById("backToChatBtn");
const backToChatFromDashboardBtn = document.getElementById(
  "backToChatFromDashboardBtn",
);
const backToChatFromAdminBtn = document.getElementById("backToChatFromAdminBtn");
const adminPage = document.getElementById("adminPage");

const accountNameEl = document.getElementById("accountName");
const accountUsernameEl = document.getElementById("accountUsername");
const accountEmailEl = document.getElementById("accountEmail");
const profileNameEl = document.getElementById("profileName");
const profileRoleEl = document.getElementById("profileRole");
const profileFocusEl = document.getElementById("profileFocus");
const supervisorModeToggleEl = document.getElementById("supervisorModeToggle");
const supervisorModeHintEl = document.getElementById("supervisorModeHint");
const stakeholderPersonaSelectEl = document.getElementById(
  "stakeholderPersonaSelect",
);
const domainPlaybookSelectEl = document.getElementById("domainPlaybookSelect");
const personaLibraryListEl = document.getElementById("personaLibraryList");
const playbookLibraryListEl = document.getElementById("playbookLibraryList");
const coachingToneEl = document.getElementById("coachingTone");
const voiceQualityEl = document.getElementById("voiceQuality");
const sessionPaceEl = document.getElementById("sessionPace");
const saveProfileBtn = document.getElementById("saveProfileBtn");
const quickPromptBtns = document.querySelectorAll("[data-quick-prompt]");

const styleDirectiveEl = document.getElementById("styleDirective");
const styleCollaborativeEl = document.getElementById("styleCollaborative");
const styleFacilitativeEl = document.getElementById("styleFacilitative");
const stylePassiveEl = document.getElementById("stylePassive");

const currentPasswordEl = document.getElementById("currentPassword");
const newPasswordEl = document.getElementById("newPassword");
const confirmPasswordEl = document.getElementById("confirmPassword");
const changePasswordBtn = document.getElementById("changePasswordBtn");

const dashboardMetaEl = document.getElementById("dashboardMeta");
const categoryScoreChartEl = document.getElementById("categoryScoreChart");
const bigFiveChartEl = document.getElementById("bigFiveChart");
const categoryTrendSvgEl = document.getElementById("categoryTrendSvg");
const trendLegendEl = document.getElementById("trendLegend");
const mbtiValueEl = document.getElementById("mbtiValue");
const discValueEl = document.getElementById("discValue");
const avgCategoryScoreEl = document.getElementById("avgCategoryScore");
const activeInterventionsEl = document.getElementById("activeInterventions");
const interventionSuccessRateEl = document.getElementById(
  "interventionSuccessRate",
);
const feedbackCoverageEl = document.getElementById("feedbackCoverage");
const feedbackLastAssessmentEl = document.getElementById("feedbackLastAssessment");
const fallbackAssessmentsEl = document.getElementById("fallbackAssessments");
const traitListEl = document.getElementById("traitList");
const strengthListEl = document.getElementById("strengthList");
const developmentListEl = document.getElementById("developmentList");
const interventionTableBodyEl = document.getElementById("interventionTableBody");
const interactionListEl = document.getElementById("interactionList");
const adminMetaEl = document.getElementById("adminMeta");
const adminTotalUsersEl = document.getElementById("adminTotalUsers");
const adminActiveAccountsEl = document.getElementById("adminActiveAccounts");
const adminInactiveAccountsEl = document.getElementById("adminInactiveAccounts");
const adminActiveUsersEl = document.getElementById("adminActiveUsers");
const adminTotalInteractionsEl = document.getElementById("adminTotalInteractions");
const adminTotalSelfReportsEl = document.getElementById("adminTotalSelfReports");
const adminTotalFeedbackEl = document.getElementById("adminTotalFeedback");
const adminAvgFeedbackEl = document.getElementById("adminAvgFeedback");
const adminDueNudgesEl = document.getElementById("adminDueNudges");
const adminTotalPracticePlansEl = document.getElementById(
  "adminTotalPracticePlans",
);
const adminSupervisorUsersEl = document.getElementById("adminSupervisorUsers");
const adminTopUseCasesEl = document.getElementById("adminTopUseCases");
const adminUsersTableBodyEl = document.getElementById("adminUsersTableBody");
const adminDisableRegistrationEl = document.getElementById(
  "adminDisableRegistration",
);
const adminSupervisorModeEl = document.getElementById("adminSupervisorMode");
const saveAdminSettingsBtn = document.getElementById("saveAdminSettingsBtn");
const adminSettingsStatusEl = document.getElementById("adminSettingsStatus");
const adminCreateUsernameEl = document.getElementById("adminCreateUsername");
const adminCreateDisplayNameEl = document.getElementById("adminCreateDisplayName");
const adminCreateEmailEl = document.getElementById("adminCreateEmail");
const adminCreateRoleEl = document.getElementById("adminCreateRole");
const adminCreatePasswordEl = document.getElementById("adminCreatePassword");
const adminCreateMustChangePasswordEl = document.getElementById(
  "adminCreateMustChangePassword",
);
const adminCreateUserBtn = document.getElementById("adminCreateUserBtn");
const adminCreateUserStatusEl = document.getElementById("adminCreateUserStatus");

const meetingNotesInputEl = document.getElementById("meetingNotesInput");
const selfReportInputEl = document.getElementById("selfReportInput");
const nudgeTitleInputEl = document.getElementById("nudgeTitleInput");
const nudgeDueAtInputEl = document.getElementById("nudgeDueAtInput");
const nudgeMessageInputEl = document.getElementById("nudgeMessageInput");
const addNudgeBtn = document.getElementById("addNudgeBtn");
const refreshNudgesBtn = document.getElementById("refreshNudgesBtn");
const nudgeStatusEl = document.getElementById("nudgeStatus");
const nudgeListEl = document.getElementById("nudgeList");
const practicePlanTitleInputEl = document.getElementById("practicePlanTitleInput");
const practicePlanObjectiveInputEl = document.getElementById(
  "practicePlanObjectiveInput",
);
const practicePlanActionsInputEl = document.getElementById(
  "practicePlanActionsInput",
);
const practicePlanCadenceSelectEl = document.getElementById(
  "practicePlanCadenceSelect",
);
const practicePlanDueAtInputEl = document.getElementById("practicePlanDueAtInput");
const practicePlanPersonaSelectEl = document.getElementById(
  "practicePlanPersonaSelect",
);
const practicePlanDomainSelectEl = document.getElementById(
  "practicePlanDomainSelect",
);
const createPracticePlanBtn = document.getElementById("createPracticePlanBtn");
const refreshPracticePlansBtn = document.getElementById("refreshPracticePlansBtn");
const practicePlanStatusEl = document.getElementById("practicePlanStatus");
const practicePlanListEl = document.getElementById("practicePlanList");
const feedbackHelpfulnessEl = document.getElementById("feedbackHelpfulness");
const feedbackClarityEl = document.getElementById("feedbackClarity");
const feedbackConfidenceEl = document.getElementById("feedbackConfidence");
const feedbackCommentEl = document.getElementById("feedbackComment");
const submitFeedbackBtn = document.getElementById("submitFeedbackBtn");
const feedbackStatusEl = document.getElementById("feedbackStatus");

const DEFAULT_STYLE_WEIGHTS = {
  directive: 2,
  collaborative: 4,
  facilitative: 4,
  passive: 2,
};
const COACHING_TONES = ["concise", "challenging", "supportive"];
const DEFAULT_COACHING_TONE = "supportive";

const CATEGORY_LABELS = {
  decisionSpeed: "Decision speed",
  delegation: "Delegation",
  conversationQuality: "Conversation quality",
  adaptability: "Adaptability",
};

const BIG_FIVE_LABELS = {
  openness: "Openness",
  conscientiousness: "Conscientiousness",
  extraversion: "Extraversion",
  agreeableness: "Agreeableness",
  neuroticism: "Neuroticism",
};

const TREND_COLORS = {
  decisionSpeed: "#0f5f9d",
  delegation: "#2f855a",
  conversationQuality: "#b7791f",
  adaptability: "#8b5cf6",
};
const VOICE_NAME_HINTS = [
  "Samantha",
  "Google UK English Female",
  "Microsoft Aria",
  "Karen",
  "Moira",
  "Daniel",
  "Alex",
];
const OPENAI_VOICE_OPTIONS = ["shimmer", "nova", "alloy", "verse", "sage"];

const state = {
  messages: [],
  listening: false,
  speaking: false,
  recognition: null,
  autoSendVoice: true,
  user: null,
  view: "chat",
  voiceQuality: "medium",
  voiceName: "shimmer",
  sessionPace: "standard",
  nudges: [],
  personas: [],
  playbooks: [],
  practicePlans: [],
  appSettings: {
    registrationDisabled: false,
    supervisorModeEnabled: true,
  },
};
let preferredVoice = null;
let lipSyncPulseTimeout = null;
let lipSyncRaf = null;
let lipSyncFallbackTimer = null;
let lipSyncDecayTimer = null;
let mouthLevel = 0;
let ttsAudioContext = null;
let ttsAudioAnalyser = null;
let ttsAudioData = null;
let ttsAudioSource = null;
let ttsAnalyserConnected = false;
let activeAudioEl = null;
let speakAbortController = null;
let speakToken = 0;
let speechRecognitionBaseText = "";
let speechRecognitionFinalText = "";
let micPermissionVerified = false;
let micStopRequested = false;
let pendingVoiceAutoSend = false;
let audioPlaybackUnlocked = false;

const STARTER_PROMPT = "";
const OPENING_MESSAGE =
  "Hi, I’m Consilium. What leadership conversation or decision would you like to work through first?";
const DRAFT_STORAGE_KEY = "consilium.messageDraft";
const MAX_CONTEXT_MESSAGES = 24;
const VOICE_CONTEXT_MESSAGES = 12;
const AUTO_SEND_MIN_CHARS = 2;

function setStatus(text) {
  statusEl.textContent = text;
}

function setCoachVisualState(mode, label = "") {
  if (!coachVisualEl || !coachVisualStatusEl) return;
  const safeMode = ["idle", "speaking", "listening"].includes(mode)
    ? mode
    : "idle";
  coachVisualEl.dataset.state = safeMode;
  if (label) {
    coachVisualStatusEl.textContent = label;
    return;
  }
  if (safeMode === "speaking") {
    coachVisualStatusEl.textContent = "Speaking";
    return;
  }
  if (safeMode === "listening") {
    coachVisualStatusEl.textContent = "Listening";
    return;
  }
  coachVisualStatusEl.textContent = "Ready to coach";
}

function refreshCoachVisualState() {
  if (state.listening) {
    setCoachVisualState("listening");
    return;
  }
  if (state.speaking) {
    setCoachVisualState("speaking");
    return;
  }
  setCoachVisualState("idle");
}

function setMouthLevel(level) {
  if (!coachAvatarEl) return;
  const clamped = Math.max(0, Math.min(1, Number(level) || 0));
  mouthLevel = clamped;
  coachAvatarEl.style.setProperty("--talk-level", clamped.toFixed(3));
  coachAvatarEl.classList.toggle("mouth-open", clamped > 0.12);
  if (coachMouthEl) {
    coachMouthEl.dataset.open = clamped > 0.12 ? "true" : "false";
  }
}

function setMouthOpen(open) {
  setMouthLevel(open ? 0.92 : 0);
}

function startFallbackLipMotion({
  min = 0.14,
  max = 0.66,
  intervalMs = 96,
} = {}) {
  if (lipSyncFallbackTimer) {
    clearInterval(lipSyncFallbackTimer);
    lipSyncFallbackTimer = null;
  }
  let phase = Math.random() * Math.PI;
  lipSyncFallbackTimer = setInterval(() => {
    phase += 0.85 + Math.random() * 0.45;
    const wave = (Math.sin(phase) + 1) / 2;
    const jitter = Math.random() * 0.22;
    const target = min + (max - min) * (wave * 0.7 + jitter * 0.3);
    setMouthLevel(target);
  }, intervalMs);
}

function pulseMouth(duration = 95, peak = 0.92) {
  if (lipSyncDecayTimer) {
    clearTimeout(lipSyncDecayTimer);
    lipSyncDecayTimer = null;
  }
  setMouthLevel(Math.max(mouthLevel, Math.max(0.1, Math.min(1, peak))));
  if (lipSyncPulseTimeout) {
    clearTimeout(lipSyncPulseTimeout);
  }
  lipSyncPulseTimeout = setTimeout(() => {
    lipSyncPulseTimeout = null;
    // If no continuous lip-sync driver is active, relax to closed mouth.
    if (!lipSyncFallbackTimer && !lipSyncRaf) {
      lipSyncDecayTimer = setTimeout(() => {
        setMouthLevel(0);
        lipSyncDecayTimer = null;
      }, Math.max(45, Math.round(duration * 0.45)));
    }
  }, Math.max(50, duration));
}

function stopLipSync() {
  if (lipSyncPulseTimeout) {
    clearTimeout(lipSyncPulseTimeout);
    lipSyncPulseTimeout = null;
  }
  if (lipSyncDecayTimer) {
    clearTimeout(lipSyncDecayTimer);
    lipSyncDecayTimer = null;
  }
  if (lipSyncFallbackTimer) {
    clearInterval(lipSyncFallbackTimer);
    lipSyncFallbackTimer = null;
  }
  if (lipSyncRaf) {
    cancelAnimationFrame(lipSyncRaf);
    lipSyncRaf = null;
  }
  setMouthLevel(0);
}

function stopAudioPlayback() {
  if (activeAudioEl) {
    activeAudioEl.pause();
    activeAudioEl.src = "";
    activeAudioEl = null;
  }
  if (speakAbortController) {
    speakAbortController.abort();
    speakAbortController = null;
  }
}

function ensureAudioAnalyzer(audioEl) {
  if (!window.AudioContext && !window.webkitAudioContext) return false;
  if (!ttsAudioContext) {
    const Ctx = window.AudioContext || window.webkitAudioContext;
    ttsAudioContext = new Ctx();
  }
  if (!ttsAudioAnalyser) {
    ttsAudioAnalyser = ttsAudioContext.createAnalyser();
    ttsAudioAnalyser.fftSize = 256;
    ttsAudioData = new Uint8Array(ttsAudioAnalyser.frequencyBinCount);
  }
  try {
    if (ttsAudioSource) {
      ttsAudioSource.disconnect();
      ttsAudioSource = null;
    }
    const source = ttsAudioContext.createMediaElementSource(audioEl);
    source.connect(ttsAudioAnalyser);
    if (!ttsAnalyserConnected) {
      ttsAudioAnalyser.connect(ttsAudioContext.destination);
      ttsAnalyserConnected = true;
    }
    ttsAudioSource = source;
  } catch {
    return false;
  }
  return true;
}

function startAudioLipSync(audioEl) {
  stopLipSync();
  const ready = ensureAudioAnalyzer(audioEl);
  if (!ready) {
    startFallbackLipMotion({ min: 0.18, max: 0.72, intervalMs: 88 });
    return;
  }

  let smoothed = 0;
  let phase = Math.random() * Math.PI;
  const tick = () => {
    if (!ttsAudioAnalyser || !ttsAudioData || !activeAudioEl) return;
    ttsAudioAnalyser.getByteFrequencyData(ttsAudioData);
    let sum = 0;
    let lowBand = 0;
    const lowCount = Math.max(1, Math.floor(ttsAudioData.length * 0.26));
    for (let i = 0; i < ttsAudioData.length; i += 1) {
      sum += ttsAudioData[i];
      if (i < lowCount) {
        lowBand += ttsAudioData[i];
      }
    }
    const avg = sum / Math.max(1, ttsAudioData.length);
    const lowAvg = lowBand / lowCount;
    const voiceEnergy = Math.max(
      0,
      Math.min(1, (avg - 8) / 38),
      Math.min(1, (lowAvg - 8) / 34),
    );
    phase += 0.23;
    const rhythmicFloor = 0.06 + ((Math.sin(phase) + 1) / 2) * 0.11;
    const target = Math.max(rhythmicFloor, voiceEnergy * 1.05);
    smoothed = smoothed * 0.67 + target * 0.33;
    setMouthLevel(smoothed);
    lipSyncRaf = requestAnimationFrame(tick);
  };
  lipSyncRaf = requestAnimationFrame(tick);
}

function setAuthStatus(text) {
  authStatusEl.textContent = text;
}

function setProfileStatus(text) {
  profileStatusEl.textContent = text;
}

function setPasswordStatus(text) {
  passwordStatusEl.textContent = text;
}

function setNudgeStatus(text) {
  if (nudgeStatusEl) {
    nudgeStatusEl.textContent = text;
  }
}

function setPracticePlanStatus(text) {
  if (practicePlanStatusEl) {
    practicePlanStatusEl.textContent = text;
  }
}

function setFeedbackStatus(text) {
  if (feedbackStatusEl) {
    feedbackStatusEl.textContent = text;
  }
}

function setAdminSettingsStatus(text) {
  if (adminSettingsStatusEl) {
    adminSettingsStatusEl.textContent = text;
  }
}

function setAdminCreateUserStatus(text) {
  if (adminCreateUserStatusEl) {
    adminCreateUserStatusEl.textContent = text;
  }
}

function normalizeAppSettings(settings) {
  const source = settings && typeof settings === "object" ? settings : {};
  return {
    registrationDisabled:
      typeof source.registrationDisabled === "boolean"
        ? source.registrationDisabled
        : false,
    supervisorModeEnabled:
      typeof source.supervisorModeEnabled === "boolean"
        ? source.supervisorModeEnabled
        : true,
  };
}

function applyAppSettings(settings) {
  state.appSettings = normalizeAppSettings(settings);
  const disabled = Boolean(state.appSettings.registrationDisabled);
  const supervisorEnabled = Boolean(state.appSettings.supervisorModeEnabled);

  if (showRegisterBtn) {
    showRegisterBtn.classList.toggle("hidden", disabled);
  }
  if (registerForm) {
    registerForm.classList.toggle("hidden", disabled);
  }
  if (disabled) {
    showAuthMode("login");
  }
  if (adminDisableRegistrationEl) {
    adminDisableRegistrationEl.checked = disabled;
  }
  if (adminSupervisorModeEl) {
    adminSupervisorModeEl.checked = supervisorEnabled;
  }
  if (supervisorModeToggleEl) {
    supervisorModeToggleEl.disabled = !supervisorEnabled;
    if (!supervisorEnabled) {
      supervisorModeToggleEl.checked = false;
    }
  }
  if (supervisorModeHintEl) {
    supervisorModeHintEl.textContent = supervisorEnabled
      ? "Supervisor mode adds more direct challenge and commitment checks."
      : "Supervisor mode is currently disabled by your administrator.";
  }
}

function normalizeStyleWeight(value, fallback) {
  const parsed = Number.parseInt(String(value), 10);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(1, Math.min(5, parsed));
}

function normalizeCoachStyleWeights(weights) {
  const source = weights && typeof weights === "object" ? weights : {};
  return {
    directive: normalizeStyleWeight(source.directive, DEFAULT_STYLE_WEIGHTS.directive),
    collaborative: normalizeStyleWeight(
      source.collaborative,
      DEFAULT_STYLE_WEIGHTS.collaborative,
    ),
    facilitative: normalizeStyleWeight(
      source.facilitative,
      DEFAULT_STYLE_WEIGHTS.facilitative,
    ),
    passive: normalizeStyleWeight(source.passive, DEFAULT_STYLE_WEIGHTS.passive),
  };
}

function applyCoachStyleWeightsToInputs(weights) {
  const normalized = normalizeCoachStyleWeights(weights);
  styleDirectiveEl.value = String(normalized.directive);
  styleCollaborativeEl.value = String(normalized.collaborative);
  styleFacilitativeEl.value = String(normalized.facilitative);
  stylePassiveEl.value = String(normalized.passive);
}

function getCoachStyleWeightsFromInputs() {
  return normalizeCoachStyleWeights({
    directive: styleDirectiveEl.value,
    collaborative: styleCollaborativeEl.value,
    facilitative: styleFacilitativeEl.value,
    passive: stylePassiveEl.value,
  });
}

function normalizeCoachingTone(value, fallback = DEFAULT_COACHING_TONE) {
  const candidate = String(value || "")
    .trim()
    .toLowerCase();
  return COACHING_TONES.includes(candidate) ? candidate : fallback;
}

function canUseCoach() {
  return Boolean(state.user && !state.user.mustChangePassword);
}

function setComposerEnabled(enabled) {
  inputEl.disabled = !enabled;
  sendBtn.disabled = !enabled;
  coachStartBtn.disabled = !enabled;
  micBtn.disabled = !enabled;
  resetBtn.disabled = !enabled;
  if (meetingNotesInputEl) meetingNotesInputEl.disabled = !enabled;
  if (selfReportInputEl) selfReportInputEl.disabled = !enabled;
}

function setView(nextView) {
  if (!state.user) return;
  const allowedViews = new Set(["chat", "settings", "dashboard"]);
  if (state.user.isAdmin) {
    allowedViews.add("admin");
  }
  state.view = allowedViews.has(nextView) ? nextView : "chat";
  const showChat = state.view === "chat";
  const showSettings = state.view === "settings";
  const showDashboard = state.view === "dashboard";
  const showAdmin = state.view === "admin";
  accountPanel.classList.toggle("hidden", !showSettings);
  dashboardPage.classList.toggle("hidden", !showDashboard);
  adminPage.classList.toggle("hidden", !showAdmin);
  chatPage.classList.toggle("hidden", !showChat);
}

function syncAccessState() {
  if (!state.user) {
    passwordRequiredBanner.classList.add("hidden");
    setComposerEnabled(false);
    setStatus("Sign in first");
    setCoachVisualState("idle", "Sign in to begin");
    return;
  }
  if (state.user.mustChangePassword) {
    passwordRequiredBanner.classList.remove("hidden");
    setComposerEnabled(false);
    setStatus("Change password to unlock coaching");
    setCoachVisualState("idle", "Change password to unlock");
    return;
  }
  passwordRequiredBanner.classList.add("hidden");
  setComposerEnabled(true);
  setStatus("Ready");
  refreshCoachVisualState();
}

function renderMessage(role, content, track = true) {
  if (track) state.messages.push({ role, content });

  const wrapper = document.createElement("article");
  wrapper.className = `msg ${role}`;

  const label = document.createElement("span");
  label.className = "msg-label";
  label.textContent = role === "assistant" ? "Consilium" : "You";

  const text = document.createElement("div");
  text.textContent = content;

  wrapper.appendChild(label);
  wrapper.appendChild(text);
  messagesEl.appendChild(wrapper);
  messagesEl.scrollTop = messagesEl.scrollHeight;
}

function formatDate(value) {
  const ms = Date.parse(String(value || ""));
  if (!Number.isFinite(ms)) return "Unknown";
  return new Date(ms).toLocaleString();
}

function toLocalDateTimeInputValue(value) {
  const ms = Date.parse(String(value || ""));
  if (!Number.isFinite(ms)) return "";
  const date = new Date(ms);
  const pad = (num) => String(num).padStart(2, "0");
  const yyyy = date.getFullYear();
  const mm = pad(date.getMonth() + 1);
  const dd = pad(date.getDate());
  const hh = pad(date.getHours());
  const min = pad(date.getMinutes());
  return `${yyyy}-${mm}-${dd}T${hh}:${min}`;
}

function formatTrend(value) {
  const numeric = Number(value || 0);
  if (!Number.isFinite(numeric) || numeric === 0) return "flat";
  return numeric > 0 ? `+${numeric}` : String(numeric);
}

function formatScore(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return "-";
  return numeric.toFixed(1);
}

function setListItems(element, items, emptyText) {
  element.innerHTML = "";
  if (!Array.isArray(items) || items.length === 0) {
    const li = document.createElement("li");
    li.className = "muted";
    li.textContent = emptyText;
    element.appendChild(li);
    return;
  }

  for (const item of items) {
    const li = document.createElement("li");
    li.textContent = String(item);
    element.appendChild(li);
  }
}

function setSelectOptions(selectEl, items, selectedValue = "", emptyLabel = "None selected") {
  if (!selectEl) return;
  const previousValue = String(selectedValue || "");
  const list = Array.isArray(items) ? items : [];
  const options = [
    `<option value="">${emptyLabel}</option>`,
    ...list.map((item) => {
      const value = String(item?.id || "");
      const label = String(item?.name || value || "Unknown");
      return `<option value="${value}">${label}</option>`;
    }),
  ];
  selectEl.innerHTML = options.join("");
  if (previousValue && list.some((item) => String(item?.id || "") === previousValue)) {
    selectEl.value = previousValue;
  } else {
    selectEl.value = "";
  }
}

function renderPersonaLibrary(personas) {
  if (!personaLibraryListEl) return;
  const list = Array.isArray(personas) ? personas : [];
  if (!list.length) {
    setListItems(personaLibraryListEl, [], "No stakeholder personas loaded yet.");
    return;
  }
  personaLibraryListEl.innerHTML = "";
  for (const persona of list) {
    const li = document.createElement("li");
    const heading = document.createElement("strong");
    heading.textContent = persona.name || "Unnamed persona";
    li.appendChild(heading);
    const detail = document.createElement("p");
    detail.className = "muted";
    const priorities = Array.isArray(persona.priorities)
      ? persona.priorities.slice(0, 3).join(", ")
      : "No priorities listed";
    detail.textContent = `${persona.archetype || "Stakeholder"} | ${persona.pressureStyle || "Pressure style not set"} | Priorities: ${priorities}`;
    li.appendChild(detail);
    personaLibraryListEl.appendChild(li);
  }
}

function renderPlaybookLibrary(playbooks) {
  if (!playbookLibraryListEl) return;
  const list = Array.isArray(playbooks) ? playbooks : [];
  if (!list.length) {
    setListItems(playbookLibraryListEl, [], "No playbooks loaded yet.");
    return;
  }
  playbookLibraryListEl.innerHTML = "";
  for (const playbook of list) {
    const li = document.createElement("li");
    const heading = document.createElement("strong");
    heading.textContent = playbook.name || "Unnamed playbook";
    li.appendChild(heading);
    const detail = document.createElement("p");
    detail.className = "muted";
    const coreMoves = Array.isArray(playbook.coreMoves)
      ? playbook.coreMoves.slice(0, 2).join(" | ")
      : "No core moves listed";
    detail.textContent = `${playbook.context || "Domain guidance"} | ${coreMoves}`;
    li.appendChild(detail);
    playbookLibraryListEl.appendChild(li);
  }
}

function renderBarChart(element, labels, scores, trends = {}) {
  element.innerHTML = "";
  for (const [key, label] of Object.entries(labels)) {
    const score = Number(scores?.[key] || 0);
    const trend = Number(trends?.[key] || 0);
    const safeScore = Number.isFinite(score) ? Math.max(1, Math.min(10, score)) : 1;

    const row = document.createElement("div");
    row.className = "bar-row";

    const heading = document.createElement("div");
    heading.className = "bar-row-head";

    const labelEl = document.createElement("span");
    labelEl.textContent = label;
    heading.appendChild(labelEl);

    const valueEl = document.createElement("span");
    valueEl.className = trend > 0 ? "trend-up" : trend < 0 ? "trend-down" : "";
    valueEl.textContent = `${safeScore.toFixed(1)} (${formatTrend(trend)})`;
    heading.appendChild(valueEl);

    const track = document.createElement("div");
    track.className = "bar-track";
    const fill = document.createElement("div");
    fill.className = "bar-fill";
    fill.style.width = `${safeScore * 10}%`;
    track.appendChild(fill);

    row.appendChild(heading);
    row.appendChild(track);
    element.appendChild(row);
  }
}

function renderTrendLegend() {
  trendLegendEl.innerHTML = "";
  for (const [key, label] of Object.entries(CATEGORY_LABELS)) {
    const item = document.createElement("span");
    item.className = "trend-key";
    const dot = document.createElement("span");
    dot.className = "trend-dot";
    dot.style.backgroundColor = TREND_COLORS[key] || "#0f5f9d";
    const text = document.createElement("span");
    text.textContent = label;
    item.appendChild(dot);
    item.appendChild(text);
    trendLegendEl.appendChild(item);
  }
}

function scoreToY(score, height, padding) {
  const clamped = Math.max(1, Math.min(10, Number(score) || 1));
  const usable = height - padding * 2;
  return padding + ((10 - clamped) / 9) * usable;
}

function buildPath(points) {
  if (points.length === 0) return "";
  let path = `M ${points[0][0].toFixed(2)} ${points[0][1].toFixed(2)}`;
  for (let i = 1; i < points.length; i += 1) {
    path += ` L ${points[i][0].toFixed(2)} ${points[i][1].toFixed(2)}`;
  }
  return path;
}

function renderTrendChart(logItems) {
  if (!categoryTrendSvgEl) return;
  categoryTrendSvgEl.innerHTML = "";

  const width = 640;
  const height = 260;
  const padding = 34;
  const chronological = Array.isArray(logItems)
    ? [...logItems].reverse().slice(-12)
    : [];

  if (chronological.length < 2) {
    categoryTrendSvgEl.innerHTML =
      '<text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" class="trend-empty">Complete at least 2 coaching interactions to see trends.</text>';
    renderTrendLegend();
    return;
  }

  const n = chronological.length;
  const xStep = (width - padding * 2) / (n - 1);
  let grid = "";
  for (let score = 2; score <= 10; score += 2) {
    const y = scoreToY(score, height, padding);
    grid += `<line x1="${padding}" y1="${y}" x2="${width - padding}" y2="${y}" class="trend-grid-line" />`;
    grid += `<text x="${padding - 8}" y="${y + 4}" text-anchor="end" class="trend-axis-label">${score}</text>`;
  }

  let paths = "";
  for (const key of Object.keys(CATEGORY_LABELS)) {
    const color = TREND_COLORS[key] || "#0f5f9d";
    const points = chronological.map((entry, idx) => {
      const x = padding + idx * xStep;
      const y = scoreToY(entry?.categoryScores?.[key], height, padding);
      return [x, y];
    });
    paths += `<path d="${buildPath(points)}" fill="none" stroke="${color}" stroke-width="2.8" stroke-linecap="round" stroke-linejoin="round" />`;
  }

  categoryTrendSvgEl.innerHTML = `
    <rect x="${padding}" y="${padding}" width="${width - padding * 2}" height="${height - padding * 2}" class="trend-bg"></rect>
    ${grid}
    ${paths}
    <line x1="${padding}" y1="${height - padding}" x2="${width - padding}" y2="${height - padding}" class="trend-axis-line" />
  `;
  renderTrendLegend();
}

function renderInterventions(interventions) {
  interventionTableBodyEl.innerHTML = "";
  if (!Array.isArray(interventions) || interventions.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 5;
    cell.className = "muted";
    cell.textContent = "No interventions recorded yet.";
    row.appendChild(cell);
    interventionTableBodyEl.appendChild(row);
    return;
  }

  for (const intervention of interventions) {
    const row = document.createElement("tr");

    const titleCell = document.createElement("td");
    titleCell.textContent = intervention.title || "Untitled";
    row.appendChild(titleCell);

    const statusCell = document.createElement("td");
    statusCell.textContent = intervention.status || "planned";
    row.appendChild(statusCell);

    const successCell = document.createElement("td");
    const success = Number(intervention.successScore);
    successCell.textContent = Number.isFinite(success) ? success.toFixed(1) : "-";
    row.appendChild(successCell);

    const targetsCell = document.createElement("td");
    targetsCell.textContent = Array.isArray(intervention.targetCategories)
      ? intervention.targetCategories.join(", ")
      : "-";
    row.appendChild(targetsCell);

    const updatedCell = document.createElement("td");
    updatedCell.textContent = formatDate(
      intervention.updatedAt || intervention.createdAt,
    );
    row.appendChild(updatedCell);

    interventionTableBodyEl.appendChild(row);
  }
}

function renderInteractions(logItems) {
  interactionListEl.innerHTML = "";
  if (!Array.isArray(logItems) || logItems.length === 0) {
    const li = document.createElement("li");
    li.className = "muted";
    li.textContent = "No assessments yet. Complete one coaching interaction.";
    interactionListEl.appendChild(li);
    return;
  }

  for (const item of logItems) {
    const li = document.createElement("li");
    const summary = document.createElement("p");
    summary.textContent = item.summary || "Assessment recorded.";
    const meta = document.createElement("p");
    meta.className = "muted";
    const evidence = [];
    if (item?.evidenceFlags?.meetingNotesProvided) evidence.push("meeting notes");
    if (item?.evidenceFlags?.selfReportProvided) evidence.push("self-report");
    const evidenceLabel = evidence.length ? ` | Evidence: ${evidence.join(", ")}` : "";
    const modeLabel =
      item?.assessmentMode === "fallback" ? "Fallback capture" : "AI assessed";
    meta.textContent = `${formatDate(item.timestamp)} | Progress: ${
      item.progress || "stable"
    } | Confidence: ${item.confidence || "medium"} | ${modeLabel}${evidenceLabel}`;
    li.appendChild(summary);
    li.appendChild(meta);
    interactionListEl.appendChild(li);
  }
}

function clearFeedbackInputs() {
  if (feedbackHelpfulnessEl) feedbackHelpfulnessEl.value = "";
  if (feedbackClarityEl) feedbackClarityEl.value = "";
  if (feedbackConfidenceEl) feedbackConfidenceEl.value = "";
  if (feedbackCommentEl) feedbackCommentEl.value = "";
}

function clearNudgeInputs() {
  if (nudgeTitleInputEl) nudgeTitleInputEl.value = "";
  if (nudgeMessageInputEl) nudgeMessageInputEl.value = "";
  if (nudgeDueAtInputEl) {
    const defaultDue = new Date(Date.now() + 48 * 60 * 60 * 1000);
    nudgeDueAtInputEl.value = toLocalDateTimeInputValue(defaultDue.toISOString());
  }
}

function normalizeNudgeStatus(value) {
  const raw = String(value || "")
    .trim()
    .toLowerCase();
  if (["scheduled", "due", "completed", "dismissed"].includes(raw)) return raw;
  return "scheduled";
}

function renderNudges(nudges) {
  if (!nudgeListEl) return;
  nudgeListEl.innerHTML = "";
  const list = Array.isArray(nudges) ? nudges : [];
  state.nudges = list;
  if (!list.length) {
    const li = document.createElement("li");
    li.className = "muted";
    li.textContent = "No nudges scheduled yet.";
    nudgeListEl.appendChild(li);
    return;
  }

  for (const nudge of list) {
    const li = document.createElement("li");
    const head = document.createElement("div");
    head.className = "nudge-item-head";

    const title = document.createElement("strong");
    title.textContent = nudge.title || "Coaching nudge";
    head.appendChild(title);

    const status = normalizeNudgeStatus(nudge.status);
    const statusEl = document.createElement("span");
    statusEl.className = `nudge-status ${status}`;
    statusEl.textContent = status;
    head.appendChild(statusEl);

    li.appendChild(head);
    const meta = document.createElement("p");
    meta.className = "muted";
    meta.textContent = `Due: ${formatDate(nudge.dueAt)}`;
    li.appendChild(meta);

    if (nudge.message) {
      const message = document.createElement("p");
      message.textContent = nudge.message;
      li.appendChild(message);
    }

    const actions = document.createElement("div");
    actions.className = "nudge-actions";

    const completeBtn = document.createElement("button");
    completeBtn.className = "btn small";
    completeBtn.type = "button";
    completeBtn.textContent = "Mark complete";
    completeBtn.disabled = status === "completed";
    completeBtn.addEventListener("click", () =>
      updateNudge(nudge.id, { status: "completed" }),
    );
    actions.appendChild(completeBtn);

    const dismissBtn = document.createElement("button");
    dismissBtn.className = "btn small";
    dismissBtn.type = "button";
    dismissBtn.textContent = "Dismiss";
    dismissBtn.disabled = status === "dismissed";
    dismissBtn.addEventListener("click", () =>
      updateNudge(nudge.id, { status: "dismissed" }),
    );
    actions.appendChild(dismissBtn);

    const reopenBtn = document.createElement("button");
    reopenBtn.className = "btn small";
    reopenBtn.type = "button";
    reopenBtn.textContent = "Reopen";
    reopenBtn.disabled = status === "scheduled" || status === "due";
    reopenBtn.addEventListener("click", () =>
      updateNudge(nudge.id, { status: "scheduled" }),
    );
    actions.appendChild(reopenBtn);

    li.appendChild(actions);
    nudgeListEl.appendChild(li);
  }
}

function normalizePracticePlanStatus(value) {
  const raw = String(value || "")
    .trim()
    .toLowerCase();
  if (["active", "paused", "completed"].includes(raw)) return raw;
  return "active";
}

function formatPracticePlanStatus(status) {
  if (status === "completed") return "completed";
  if (status === "paused") return "dismissed";
  return "scheduled";
}

function clearPracticePlanInputs() {
  if (practicePlanTitleInputEl) practicePlanTitleInputEl.value = "";
  if (practicePlanObjectiveInputEl) practicePlanObjectiveInputEl.value = "";
  if (practicePlanActionsInputEl) practicePlanActionsInputEl.value = "";
  if (practicePlanCadenceSelectEl) practicePlanCadenceSelectEl.value = "weekly";
  if (practicePlanDueAtInputEl) practicePlanDueAtInputEl.value = "";
  if (practicePlanPersonaSelectEl) practicePlanPersonaSelectEl.value = "";
  if (practicePlanDomainSelectEl) practicePlanDomainSelectEl.value = "";
}

function renderPracticePlans(plans) {
  if (!practicePlanListEl) return;
  const list = Array.isArray(plans) ? plans : [];
  state.practicePlans = list;
  practicePlanListEl.innerHTML = "";
  if (!list.length) {
    const li = document.createElement("li");
    li.className = "muted";
    li.textContent = "No practice plans yet.";
    practicePlanListEl.appendChild(li);
    return;
  }

  for (const plan of list) {
    const li = document.createElement("li");
    const head = document.createElement("div");
    head.className = "practice-item-head";

    const title = document.createElement("strong");
    title.textContent = plan.title || "Practice plan";
    head.appendChild(title);

    const status = normalizePracticePlanStatus(plan.status);
    const statusEl = document.createElement("span");
    statusEl.className = `nudge-status ${formatPracticePlanStatus(status)}`;
    statusEl.textContent = status;
    head.appendChild(statusEl);
    li.appendChild(head);

    const meta = document.createElement("p");
    meta.className = "muted practice-meta";
    const cadence = plan.cadence ? `Cadence: ${plan.cadence}` : "Cadence: not set";
    const dueAt = plan.nextDueAt ? `Next due: ${formatDate(plan.nextDueAt)}` : "No due date";
    const completedCount = Number(plan.completedCount || 0);
    meta.textContent = `${cadence} | ${dueAt} | Sessions completed: ${completedCount}`;
    li.appendChild(meta);

    const context = [];
    if (plan.stakeholderPersona?.name) {
      context.push(`Persona: ${plan.stakeholderPersona.name}`);
    }
    if (plan.playbook?.name) {
      context.push(`Playbook: ${plan.playbook.name}`);
    }
    if (context.length) {
      const contextEl = document.createElement("p");
      contextEl.className = "muted";
      contextEl.textContent = context.join(" | ");
      li.appendChild(contextEl);
    }

    if (plan.objective) {
      const objective = document.createElement("p");
      objective.textContent = plan.objective;
      li.appendChild(objective);
    }

    if (Array.isArray(plan.actions) && plan.actions.length) {
      const actionList = document.createElement("ul");
      actionList.className = "dash-bullets";
      for (const action of plan.actions) {
        const item = document.createElement("li");
        item.textContent = action;
        actionList.appendChild(item);
      }
      li.appendChild(actionList);
    }

    const actions = document.createElement("div");
    actions.className = "nudge-actions";

    const practicedBtn = document.createElement("button");
    practicedBtn.className = "btn small";
    practicedBtn.type = "button";
    practicedBtn.textContent = "Mark practiced";
    practicedBtn.disabled = status === "completed";
    practicedBtn.addEventListener("click", () =>
      updatePracticePlan(plan.id, { action: "mark_practiced" }),
    );
    actions.appendChild(practicedBtn);

    const toggleBtn = document.createElement("button");
    toggleBtn.className = "btn small";
    toggleBtn.type = "button";
    toggleBtn.textContent = status === "paused" ? "Activate" : "Pause";
    toggleBtn.disabled = status === "completed";
    toggleBtn.addEventListener("click", () =>
      updatePracticePlan(plan.id, {
        action: status === "paused" ? "activate" : "pause",
      }),
    );
    actions.appendChild(toggleBtn);

    const completeBtn = document.createElement("button");
    completeBtn.className = "btn small";
    completeBtn.type = "button";
    completeBtn.textContent = "Complete";
    completeBtn.disabled = status === "completed";
    completeBtn.addEventListener("click", () =>
      updatePracticePlan(plan.id, { action: "complete" }),
    );
    actions.appendChild(completeBtn);

    li.appendChild(actions);
    practicePlanListEl.appendChild(li);
  }
}

function renderDashboard(dashboard) {
  const profile = dashboard?.profile || {};
  const metrics = profile.metrics || {};
  const retentionDays = Number(dashboard?.retentionDays || 365);
  const generatedAt = formatDate(dashboard?.generatedAt);
  dashboardMetaEl.textContent = `Generated ${generatedAt}. Retention: ${retentionDays} days. Last profile update: ${formatDate(
    profile.updatedAt,
  )}.`;

  renderBarChart(
    categoryScoreChartEl,
    CATEGORY_LABELS,
    profile.categoryScores || {},
    profile.categoryTrends || {},
  );
  renderBarChart(
    bigFiveChartEl,
    BIG_FIVE_LABELS,
    profile?.personality?.bigFive || {},
    {},
  );
  renderTrendChart(profile.interactionLog || []);

  mbtiValueEl.textContent = profile?.personality?.mbti || "Unknown";
  discValueEl.textContent = profile?.personality?.disc || "Unknown";
  avgCategoryScoreEl.textContent = formatScore(metrics.averageCategoryScore);
  activeInterventionsEl.textContent = String(metrics.activeInterventions || 0);
  interventionSuccessRateEl.textContent = `${formatScore(
    metrics.interventionSuccessRate,
  )}%`;

  const loop = profile?.feedbackLoop || {};
  const totalInteractions = Number(loop.totalInteractions || 0);
  const analyzedInteractions = Number(loop.analyzedInteractions || 0);
  const fallbackInteractions = Number(loop.fallbackInteractions || 0);
  if (feedbackCoverageEl) {
    feedbackCoverageEl.textContent = `${analyzedInteractions}/${totalInteractions}`;
  }
  if (fallbackAssessmentsEl) {
    fallbackAssessmentsEl.textContent = String(fallbackInteractions);
  }
  if (feedbackLastAssessmentEl) {
    feedbackLastAssessmentEl.textContent = loop.lastAssessmentAt
      ? formatDate(loop.lastAssessmentAt)
      : "Not yet";
  }

  setListItems(traitListEl, profile.traits, "No traits identified yet.");
  setListItems(strengthListEl, profile.strengths, "No strengths captured yet.");
  setListItems(
    developmentListEl,
    profile.developmentAreas,
    "No development areas captured yet.",
  );
  renderInterventions(profile.interventions || []);
  renderInteractions(profile.interactionLog || []);
}

async function showDashboard() {
  if (!state.user) return;

  try {
    const payload = await apiRequest("/api/dashboard", { method: "GET" });
    if (payload.user) {
      applyUserToUI(payload.user);
    }
    renderDashboard(payload.dashboard || {});
    setView("dashboard");
  } catch (err) {
    setStatus("Dashboard error");
    renderMessage("assistant", `Error: ${err.message}`, false);
  }
}

function formatUseCaseList(items, limit = 3) {
  if (!Array.isArray(items) || items.length === 0) return "General coaching";
  return items
    .slice(0, limit)
    .map((item) => `${item.label} (${item.count})`)
    .join(", ");
}

function renderAdminDashboard(dashboard) {
  const summary = dashboard?.summary || {};
  const users = Array.isArray(dashboard?.users) ? dashboard.users : [];
  const generatedAt = formatDate(dashboard?.generatedAt);
  if (adminMetaEl) {
    adminMetaEl.textContent = `Generated ${generatedAt}. Admin-only view across all accounts.`;
  }
  if (adminTotalUsersEl) {
    adminTotalUsersEl.textContent = String(summary.totalUsers || 0);
  }
  if (adminActiveAccountsEl) {
    adminActiveAccountsEl.textContent = String(summary.activeAccounts || 0);
  }
  if (adminInactiveAccountsEl) {
    adminInactiveAccountsEl.textContent = String(summary.inactiveAccounts || 0);
  }
  if (adminActiveUsersEl) {
    adminActiveUsersEl.textContent = String(summary.activeUsers30d || 0);
  }
  if (adminTotalInteractionsEl) {
    adminTotalInteractionsEl.textContent = String(summary.totalInteractions || 0);
  }
  if (adminTotalSelfReportsEl) {
    adminTotalSelfReportsEl.textContent = String(summary.totalSelfReports || 0);
  }
  if (adminTotalFeedbackEl) {
    adminTotalFeedbackEl.textContent = String(summary.totalFeedbackSubmissions || 0);
  }
  if (adminAvgFeedbackEl) {
    adminAvgFeedbackEl.textContent = Number.isFinite(summary.averageFeedbackScore)
      ? formatScore(summary.averageFeedbackScore)
      : "-";
  }
  if (adminDueNudgesEl) {
    adminDueNudgesEl.textContent = String(summary.totalNudgesDue || 0);
  }
  if (adminTotalPracticePlansEl) {
    adminTotalPracticePlansEl.textContent = String(summary.totalPracticePlans || 0);
  }
  if (adminSupervisorUsersEl) {
    adminSupervisorUsersEl.textContent = String(summary.usersInSupervisorMode || 0);
  }
  if (adminDisableRegistrationEl) {
    adminDisableRegistrationEl.checked = Boolean(summary.registrationDisabled);
  }
  if (adminSupervisorModeEl) {
    adminSupervisorModeEl.checked = Boolean(summary.supervisorModeEnabled);
  }
  if (adminTopUseCasesEl) {
    const labels = Array.isArray(summary.topUseCases)
      ? summary.topUseCases.map((item) => `${item.label}: ${item.count}`)
      : [];
    setListItems(adminTopUseCasesEl, labels, "No use-case data yet.");
  }

  if (!adminUsersTableBodyEl) return;
  adminUsersTableBodyEl.innerHTML = "";

  if (!users.length) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 9;
    cell.className = "muted";
    cell.textContent = "No users found.";
    row.appendChild(cell);
    adminUsersTableBodyEl.appendChild(row);
    return;
  }

  for (const user of users) {
    const row = document.createElement("tr");

    const userCell = document.createElement("td");
    userCell.textContent = `${user.displayName || user.username || "Unknown"} (@${
      user.username || "-"
    })`;
    const userMeta = document.createElement("p");
    userMeta.className = "muted";
    userMeta.textContent = `${user.role || "No role"} | Created ${formatDate(
      user.createdAt,
    )}`;
    userCell.appendChild(userMeta);
    row.appendChild(userCell);

    const statusCell = document.createElement("td");
    const statusPill = document.createElement("span");
    statusPill.className = `status-pill ${user.isActive ? "active" : "inactive"}`;
    statusPill.textContent = user.isActive ? "Active" : "Inactive";
    statusCell.appendChild(statusPill);
    const statusMeta = document.createElement("p");
    statusMeta.className = "muted";
    const nudgeSummary = user.nudgeSummary || {};
    const practiceSummary = user.practicePlansSummary || {};
    statusMeta.textContent = `Must change password: ${
      user.mustChangePassword ? "Yes" : "No"
    } | Supervisor mode: ${user.supervisorMode ? "On" : "Off"} | Nudges due: ${
      nudgeSummary.due || 0
    } | Active plans: ${practiceSummary.active || 0}`;
    statusCell.appendChild(statusMeta);
    row.appendChild(statusCell);

    const usageCell = document.createElement("td");
    const usage = user.usage || {};
    usageCell.textContent = `${usage.totalInteractions || 0} interactions`;
    const usageMeta = document.createElement("p");
    usageMeta.className = "muted";
    usageMeta.textContent = `Busy ${
      usage?.sessionPace?.busy || 0
    } | Standard ${usage?.sessionPace?.standard || 0} | Fallback ${
      usage.fallbackAssessments || 0
    }`;
    usageCell.appendChild(usageMeta);
    row.appendChild(usageCell);

    const usedForCell = document.createElement("td");
    usedForCell.textContent = formatUseCaseList(user.usedFor, 3);
    const usedForMeta = document.createElement("p");
    usedForMeta.className = "muted";
    const personaName = user?.selectedStakeholderPersona?.name || "None";
    const playbookName = user?.selectedPlaybook?.name || "None";
    usedForMeta.textContent = `Focus: ${
      user.focusAreas || "Not set"
    } | Persona: ${personaName} | Playbook: ${playbookName}`;
    usedForCell.appendChild(usedForMeta);
    row.appendChild(usedForCell);

    const personalityCell = document.createElement("td");
    const personality = user.personality || {};
    const bigFive = personality.bigFive || {};
    personalityCell.textContent = `MBTI ${personality.mbti || "Unknown"} | DISC ${
      personality.disc || "Unknown"
    }`;
    const personalityMeta = document.createElement("p");
    personalityMeta.className = "muted";
    personalityMeta.textContent = `O:${bigFive.openness ?? "-"} C:${
      bigFive.conscientiousness ?? "-"
    } E:${bigFive.extraversion ?? "-"} A:${bigFive.agreeableness ?? "-"} N:${
      bigFive.neuroticism ?? "-"
    }`;
    personalityCell.appendChild(personalityMeta);
    row.appendChild(personalityCell);

    const scoreCell = document.createElement("td");
    const scores = user.categoryScores || {};
    const metrics = user.metrics || {};
    scoreCell.textContent = `Avg ${formatScore(metrics.averageCategoryScore)}`;
    const scoreMeta = document.createElement("p");
    scoreMeta.className = "muted";
    scoreMeta.textContent = `D:${scores.decisionSpeed ?? "-"} Del:${
      scores.delegation ?? "-"
    } C:${scores.conversationQuality ?? "-"} A:${scores.adaptability ?? "-"}`;
    scoreCell.appendChild(scoreMeta);
    row.appendChild(scoreCell);

    const feedbackCell = document.createElement("td");
    const feedback = user.feedback || {};
    const coacheeFeedback = Array.isArray(feedback.recentCoacheeFeedback)
      ? feedback.recentCoacheeFeedback
      : [];
    const selfReports = Array.isArray(feedback.recentSelfReports)
      ? feedback.recentSelfReports
      : [];
    if (coacheeFeedback.length) {
      const latest = coacheeFeedback[0] || {};
      const scores = [latest.helpfulness, latest.clarity, latest.confidence]
        .filter((value) => Number.isFinite(value))
        .join("/");
      feedbackCell.textContent = latest.comment || `Scores: ${scores || "-"}`;
      const feedbackMeta = document.createElement("p");
      feedbackMeta.className = "muted";
      feedbackMeta.textContent = `${formatDate(latest.timestamp)} | Ratings: ${
        usage.feedbackSubmissionCount || 0
      } | Avg ${Number.isFinite(usage.averageFeedbackScore) ? formatScore(usage.averageFeedbackScore) : "-"}`;
      feedbackCell.appendChild(feedbackMeta);
    } else if (!selfReports.length) {
      feedbackCell.textContent = "No feedback yet";
      const feedbackMeta = document.createElement("p");
      feedbackMeta.className = "muted";
      feedbackMeta.textContent = `Self-reports: ${usage.selfReportCount || 0} | Ratings: ${
        usage.feedbackSubmissionCount || 0
      }`;
      feedbackCell.appendChild(feedbackMeta);
    } else {
      feedbackCell.textContent = selfReports[0]?.text || "Feedback captured";
      const feedbackMeta = document.createElement("p");
      feedbackMeta.className = "muted";
      feedbackMeta.textContent = `${formatDate(
        selfReports[0]?.timestamp,
      )} | Self-reports: ${usage.selfReportCount || 0} | Ratings: ${
        usage.feedbackSubmissionCount || 0
      }`;
      feedbackCell.appendChild(feedbackMeta);
    }
    row.appendChild(feedbackCell);

    const lastCell = document.createElement("td");
    lastCell.textContent = formatDate(user?.usage?.lastInteractionAt);
    row.appendChild(lastCell);

    const actionsCell = document.createElement("td");
    const actions = document.createElement("div");
    actions.className = "admin-actions";

    const toggleActiveBtn = document.createElement("button");
    toggleActiveBtn.className = "btn small";
    toggleActiveBtn.type = "button";
    toggleActiveBtn.textContent = user.isActive ? "Deactivate" : "Activate";
    toggleActiveBtn.disabled = user.id === state.user?.id && user.isActive;
    toggleActiveBtn.addEventListener("click", () =>
      runAdminUserAction(user.id, user.isActive ? "deactivate" : "activate"),
    );
    actions.appendChild(toggleActiveBtn);

    const forceResetBtn = document.createElement("button");
    forceResetBtn.className = "btn small";
    forceResetBtn.type = "button";
    forceResetBtn.textContent = "Force reset";
    forceResetBtn.addEventListener("click", () =>
      runAdminUserAction(user.id, "force_password_reset"),
    );
    actions.appendChild(forceResetBtn);

    const tempPasswordBtn = document.createElement("button");
    tempPasswordBtn.className = "btn small";
    tempPasswordBtn.type = "button";
    tempPasswordBtn.textContent = "Set temp password";
    tempPasswordBtn.addEventListener("click", () => {
      const tempPassword = window.prompt(
        `Set temporary password for @${user.username} (min 10 chars):`,
      );
      if (!tempPassword) return;
      runAdminUserAction(user.id, "reset_password", { newPassword: tempPassword });
    });
    actions.appendChild(tempPasswordBtn);

    actionsCell.appendChild(actions);
    row.appendChild(actionsCell);

    adminUsersTableBodyEl.appendChild(row);
  }
}

async function runAdminUserAction(userId, action, extra = {}) {
  if (!state.user || !state.user.isAdmin) return;
  setStatus("Applying admin action...");
  try {
    const payload = await apiRequest(
      `/api/admin/users/${encodeURIComponent(userId)}/action`,
      {
        method: "POST",
        body: JSON.stringify({
          action,
          ...extra,
        }),
      },
    );
    if (payload.settings) {
      applyAppSettings(payload.settings);
    }
    if (payload.user) {
      applyUserToUI(payload.user);
    }
    if (payload.dashboard) {
      renderAdminDashboard(payload.dashboard);
    } else {
      await showAdminDashboard();
    }
    setStatus("Admin action applied.");
  } catch (err) {
    setStatus(`Admin action failed: ${err.message}`);
  }
}

async function saveAdminSettings() {
  if (!state.user || !state.user.isAdmin) return;
  const registrationDisabled = Boolean(adminDisableRegistrationEl?.checked);
  const supervisorModeEnabled = Boolean(adminSupervisorModeEl?.checked);
  setAdminSettingsStatus("Saving...");
  if (saveAdminSettingsBtn) saveAdminSettingsBtn.disabled = true;
  try {
    const payload = await apiRequest("/api/admin/settings", {
      method: "PUT",
      body: JSON.stringify({ registrationDisabled, supervisorModeEnabled }),
    });
    if (payload.settings) {
      applyAppSettings(payload.settings);
    }
    if (payload.user) {
      applyUserToUI(payload.user);
    }
    if (payload.dashboard) {
      renderAdminDashboard(payload.dashboard);
    }
    setAdminSettingsStatus("Admin settings saved.");
  } catch (err) {
    setAdminSettingsStatus(err.message);
  } finally {
    if (saveAdminSettingsBtn) saveAdminSettingsBtn.disabled = false;
  }
}

function clearAdminCreateUserForm() {
  if (adminCreateUsernameEl) adminCreateUsernameEl.value = "";
  if (adminCreateDisplayNameEl) adminCreateDisplayNameEl.value = "";
  if (adminCreateEmailEl) adminCreateEmailEl.value = "";
  if (adminCreateRoleEl) adminCreateRoleEl.value = "";
  if (adminCreatePasswordEl) adminCreatePasswordEl.value = "";
  if (adminCreateMustChangePasswordEl) adminCreateMustChangePasswordEl.checked = true;
}

async function createAdminUser() {
  if (!state.user || !state.user.isAdmin) return;

  const username = String(adminCreateUsernameEl?.value || "").trim();
  const displayName = String(adminCreateDisplayNameEl?.value || "").trim();
  const email = String(adminCreateEmailEl?.value || "").trim();
  const role = String(adminCreateRoleEl?.value || "").trim();
  const password = String(adminCreatePasswordEl?.value || "");
  const mustChangePassword = Boolean(adminCreateMustChangePasswordEl?.checked);

  if (!username) {
    setAdminCreateUserStatus("Username is required.");
    return;
  }
  if (!password || password.length < 8) {
    setAdminCreateUserStatus("Temporary password must be at least 8 characters.");
    return;
  }

  setAdminCreateUserStatus("Creating user...");
  if (adminCreateUserBtn) adminCreateUserBtn.disabled = true;

  try {
    const payload = await apiRequest("/api/admin/users", {
      method: "POST",
      body: JSON.stringify({
        username,
        displayName,
        email,
        role,
        password,
        mustChangePassword,
      }),
    });
    if (payload.settings) {
      applyAppSettings(payload.settings);
    }
    if (payload.user) {
      applyUserToUI(payload.user);
    }
    if (payload.dashboard) {
      renderAdminDashboard(payload.dashboard);
    }
    clearAdminCreateUserForm();
    const createdName =
      payload?.createdUser?.username || payload?.createdUser?.displayName || username;
    setAdminCreateUserStatus(`Created @${createdName}.`);
  } catch (err) {
    setAdminCreateUserStatus(err.message);
  } finally {
    if (adminCreateUserBtn) adminCreateUserBtn.disabled = false;
  }
}

async function showAdminDashboard() {
  if (!state.user || !state.user.isAdmin) return;
  try {
    const payload = await apiRequest("/api/admin/dashboard", { method: "GET" });
    if (payload.settings) {
      applyAppSettings(payload.settings);
    }
    if (payload.user) {
      applyUserToUI(payload.user);
    }
    renderAdminDashboard(payload.dashboard || {});
    setView("admin");
  } catch (err) {
    setStatus("Admin dashboard error");
    renderMessage("assistant", `Error: ${err.message}`, false);
  }
}

function buildFeedbackPayload() {
  const toScore = (element) => {
    const raw = element ? Number.parseInt(String(element.value || "").trim(), 10) : NaN;
    if (!Number.isFinite(raw)) return null;
    return Math.max(1, Math.min(10, raw));
  };
  return {
    helpfulness: toScore(feedbackHelpfulnessEl),
    clarity: toScore(feedbackClarityEl),
    confidence: toScore(feedbackConfidenceEl),
    comment: feedbackCommentEl ? feedbackCommentEl.value.trim() : "",
  };
}

function hasFeedbackValue(payload) {
  if (!payload || typeof payload !== "object") return false;
  return (
    Number.isFinite(payload.helpfulness) ||
    Number.isFinite(payload.clarity) ||
    Number.isFinite(payload.confidence) ||
    Boolean(payload.comment)
  );
}

async function submitSessionFeedback() {
  if (!state.user) return;
  const payload = buildFeedbackPayload();
  if (!hasFeedbackValue(payload)) {
    setFeedbackStatus("Add at least one score or a comment.");
    return;
  }
  setFeedbackStatus("Submitting...");
  if (submitFeedbackBtn) submitFeedbackBtn.disabled = true;

  try {
    const response = await apiRequest("/api/feedback", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    if (response.user) {
      applyUserToUI(response.user);
    }
    if (response.dashboard && state.view === "dashboard") {
      renderDashboard(response.dashboard);
    }
    clearFeedbackInputs();
    setFeedbackStatus("Feedback saved.");
  } catch (err) {
    setFeedbackStatus(err.message);
  } finally {
    if (submitFeedbackBtn) submitFeedbackBtn.disabled = false;
  }
}

async function loadNudges({ quiet = false } = {}) {
  if (!state.user) return;
  if (!quiet) setNudgeStatus("Loading nudges...");
  try {
    const payload = await apiRequest("/api/nudges", { method: "GET" });
    const list = Array.isArray(payload.nudges) ? payload.nudges : [];
    renderNudges(list);
    const summary = payload.summary || {};
    const due = Number(summary.due || 0);
    if (!quiet) {
      setNudgeStatus(
        due > 0
          ? `${due} nudge${due === 1 ? "" : "s"} due now.`
          : "Nudges updated.",
      );
    }
  } catch (err) {
    if (!quiet) setNudgeStatus(err.message);
  }
}

async function createNudge() {
  if (!state.user) return;
  const title = nudgeTitleInputEl ? nudgeTitleInputEl.value.trim() : "";
  const message = nudgeMessageInputEl ? nudgeMessageInputEl.value.trim() : "";
  const dueAtRaw = nudgeDueAtInputEl ? nudgeDueAtInputEl.value : "";
  if (!title && !message) {
    setNudgeStatus("Add a title or message for the nudge.");
    return;
  }
  setNudgeStatus("Scheduling...");
  if (addNudgeBtn) addNudgeBtn.disabled = true;

  try {
    const payload = await apiRequest("/api/nudges", {
      method: "POST",
      body: JSON.stringify({
        title,
        message,
        dueAt: dueAtRaw ? new Date(dueAtRaw).toISOString() : "",
      }),
    });
    if (payload.user) {
      applyUserToUI(payload.user);
    }
    renderNudges(Array.isArray(payload.nudges) ? payload.nudges : []);
    clearNudgeInputs();
    setNudgeStatus("Nudge scheduled.");
  } catch (err) {
    setNudgeStatus(err.message);
  } finally {
    if (addNudgeBtn) addNudgeBtn.disabled = false;
  }
}

async function updateNudge(nudgeId, updates) {
  if (!state.user) return;
  setNudgeStatus("Updating nudge...");
  try {
    const payload = await apiRequest(`/api/nudges/${encodeURIComponent(nudgeId)}`, {
      method: "PUT",
      body: JSON.stringify(updates || {}),
    });
    if (payload.user) {
      applyUserToUI(payload.user);
    }
    renderNudges(Array.isArray(payload.nudges) ? payload.nudges : []);
    setNudgeStatus("Nudge updated.");
  } catch (err) {
    setNudgeStatus(err.message);
  }
}

function applyPersonaAndPlaybookSelectors(user = null) {
  const selectedPersona = String(user?.selectedStakeholderPersonaId || "");
  const selectedPlaybook = String(user?.selectedPlaybookId || "");
  setSelectOptions(
    stakeholderPersonaSelectEl,
    state.personas,
    selectedPersona,
    "None selected",
  );
  setSelectOptions(
    domainPlaybookSelectEl,
    state.playbooks,
    selectedPlaybook,
    "None selected",
  );

  const practicePersona =
    String(practicePlanPersonaSelectEl?.value || selectedPersona || "");
  const practiceDomain =
    String(practicePlanDomainSelectEl?.value || selectedPlaybook || "");
  setSelectOptions(
    practicePlanPersonaSelectEl,
    state.personas,
    practicePersona,
    "None selected",
  );
  setSelectOptions(
    practicePlanDomainSelectEl,
    state.playbooks,
    practiceDomain,
    "None selected",
  );
}

async function loadCoachLibraries({ quiet = false } = {}) {
  if (!state.user) return;
  try {
    const [personaPayload, playbookPayload] = await Promise.all([
      apiRequest("/api/personas", { method: "GET" }),
      apiRequest("/api/playbooks", { method: "GET" }),
    ]);
    state.personas = Array.isArray(personaPayload.personas)
      ? personaPayload.personas
      : [];
    state.playbooks = Array.isArray(playbookPayload.playbooks)
      ? playbookPayload.playbooks
      : [];
    renderPersonaLibrary(state.personas);
    renderPlaybookLibrary(state.playbooks);
    applyPersonaAndPlaybookSelectors(state.user);
  } catch (err) {
    if (!quiet) {
      setProfileStatus(`Library load failed: ${err.message}`);
    }
  }
}

async function loadPracticePlans({ quiet = false } = {}) {
  if (!state.user) return;
  if (!quiet) setPracticePlanStatus("Loading practice plans...");
  try {
    const payload = await apiRequest("/api/practice-plans", { method: "GET" });
    renderPracticePlans(Array.isArray(payload.plans) ? payload.plans : []);
    if (!quiet) {
      const summary = payload.summary || {};
      setPracticePlanStatus(
        `${summary.active || 0} active, ${summary.completed || 0} completed.`,
      );
    }
  } catch (err) {
    if (!quiet) setPracticePlanStatus(err.message);
  }
}

async function createPracticePlan() {
  if (!state.user) return;
  const title = String(practicePlanTitleInputEl?.value || "").trim();
  const objective = String(practicePlanObjectiveInputEl?.value || "").trim();
  const actions = String(practicePlanActionsInputEl?.value || "").trim();
  const cadence = String(practicePlanCadenceSelectEl?.value || "weekly");
  const stakeholderPersonaId = String(practicePlanPersonaSelectEl?.value || "");
  const domainId = String(practicePlanDomainSelectEl?.value || "");
  const dueAtRaw = String(practicePlanDueAtInputEl?.value || "");

  if (!title) {
    setPracticePlanStatus("Plan title is required.");
    return;
  }

  setPracticePlanStatus("Creating practice plan...");
  if (createPracticePlanBtn) createPracticePlanBtn.disabled = true;
  try {
    const payload = await apiRequest("/api/practice-plans", {
      method: "POST",
      body: JSON.stringify({
        title,
        objective,
        actions,
        cadence,
        stakeholderPersonaId,
        domainId,
        nextDueAt: dueAtRaw ? new Date(dueAtRaw).toISOString() : "",
      }),
    });
    if (payload.user) {
      applyUserToUI(payload.user);
    }
    renderPracticePlans(Array.isArray(payload.plans) ? payload.plans : []);
    clearPracticePlanInputs();
    setPracticePlanStatus("Practice plan created.");
  } catch (err) {
    setPracticePlanStatus(err.message);
  } finally {
    if (createPracticePlanBtn) createPracticePlanBtn.disabled = false;
  }
}

async function updatePracticePlan(planId, updates) {
  if (!state.user) return;
  setPracticePlanStatus("Updating practice plan...");
  try {
    const payload = await apiRequest(
      `/api/practice-plans/${encodeURIComponent(planId)}`,
      {
        method: "PUT",
        body: JSON.stringify(updates || {}),
      },
    );
    if (payload.user) {
      applyUserToUI(payload.user);
    }
    renderPracticePlans(Array.isArray(payload.plans) ? payload.plans : []);
    setPracticePlanStatus("Practice plan updated.");
  } catch (err) {
    setPracticePlanStatus(err.message);
  }
}

function normalizeSpeechText(text) {
  return String(text || "")
    .replace(/\*\*/g, "")
    .replace(/`/g, "")
    .replace(/^#{1,6}\s+/gm, "")
    .replace(/^\s*[-*]\s+/gm, "")
    .replace(/^\s*\d+\.\s+/gm, "")
    .replace(/\s+/g, " ")
    .trim();
}

function splitSpeechChunks(text, maxChars = 220) {
  const chunks = [];
  const sentences = text.split(/(?<=[.!?])\s+/);
  let current = "";
  for (const sentence of sentences) {
    if (!sentence) continue;
    if (!current) {
      current = sentence;
      continue;
    }
    if ((current + " " + sentence).length <= maxChars) {
      current += ` ${sentence}`;
    } else {
      chunks.push(current);
      current = sentence;
    }
  }
  if (current) chunks.push(current);
  return chunks.length ? chunks : [text];
}

function normalizeSpeakerName(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

function parseRolePlaySegments(text) {
  const raw = String(text || "");
  const lines = raw
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  if (!lines.length) return [];

  const segments = [];
  for (const line of lines) {
    const match = line.match(/^(?:[-*]\s*)?([A-Za-z][A-Za-z0-9 '&/.-]{1,32}):\s*(.+)$/);
    if (!match) continue;
    const speaker = normalizeSpeakerName(match[1]);
    const content = String(match[2] || "").trim();
    if (!speaker || !content) continue;
    segments.push({ speaker, text: content });
  }

  const uniqueSpeakers = new Set(
    segments.map((segment) => normalizeSpeakerName(segment.speaker).toLowerCase()),
  );
  if (segments.length < 2 || uniqueSpeakers.size < 2) {
    // Fallback for single-line role-play text containing multiple speaker labels.
    const inlineMatches = [];
    const regex = /([A-Za-z][A-Za-z0-9 '&/.-]{1,32}):\s*/g;
    let match = regex.exec(raw);
    while (match) {
      const speaker = normalizeSpeakerName(match[1]);
      const contentStart = regex.lastIndex;
      const nextMatch = regex.exec(raw);
      const contentEnd = nextMatch ? nextMatch.index : raw.length;
      const content = raw.slice(contentStart, contentEnd).trim();
      if (speaker && content) {
        inlineMatches.push({ speaker, text: content });
      }
      if (!nextMatch) break;
      regex.lastIndex = nextMatch.index;
      match = regex.exec(raw);
    }
    const inlineSpeakers = new Set(
      inlineMatches.map((item) => normalizeSpeakerName(item.speaker).toLowerCase()),
    );
    if (inlineMatches.length >= 2 && inlineSpeakers.size >= 2) {
      return inlineMatches;
    }
    return [];
  }
  return segments;
}

function isCoachSpeaker(speaker) {
  const key = normalizeSpeakerName(speaker).toLowerCase();
  if (!key) return true;
  return (
    key.includes("consilium") ||
    key.includes("coach") ||
    key.includes("facilitator") ||
    key.includes("narrator") ||
    key === "assistant"
  );
}

function buildOpenAIVoiceAssignments(segments, primaryVoice) {
  const primary = String(primaryVoice || "shimmer").toLowerCase();
  const pool = OPENAI_VOICE_OPTIONS.filter((voice) => voice !== primary);
  const assignments = new Map();
  let idx = 0;

  for (const segment of segments) {
    const key = normalizeSpeakerName(segment.speaker);
    if (!key || assignments.has(key)) continue;
    if (isCoachSpeaker(key)) {
      assignments.set(key, primary);
      continue;
    }
    const fallbackVoice = OPENAI_VOICE_OPTIONS[0] || primary;
    const assigned = pool.length ? pool[idx % pool.length] : fallbackVoice;
    assignments.set(key, assigned);
    idx += 1;
  }
  return assignments;
}

function buildBrowserVoiceAssignments(segments, coachVoice) {
  const englishVoices = (window.speechSynthesis?.getVoices?.() || []).filter((voice) =>
    String(voice.lang || "")
      .toLowerCase()
      .startsWith("en"),
  );
  const pool = englishVoices.filter((voice) => !coachVoice || voice.voiceURI !== coachVoice.voiceURI);
  const assignments = new Map();
  let idx = 0;

  for (const segment of segments) {
    const key = normalizeSpeakerName(segment.speaker);
    if (!key || assignments.has(key)) continue;
    if (isCoachSpeaker(key)) {
      assignments.set(key, coachVoice || null);
      continue;
    }
    const assigned = pool.length ? pool[idx % pool.length] : coachVoice || null;
    assignments.set(key, assigned);
    idx += 1;
  }

  return assignments;
}

function pickPreferredVoice() {
  if (!("speechSynthesis" in window)) return null;
  const voices = window.speechSynthesis.getVoices();
  if (!voices.length) return null;

  if (preferredVoice && voices.some((v) => v.voiceURI === preferredVoice.voiceURI)) {
    return preferredVoice;
  }

  for (const hint of VOICE_NAME_HINTS) {
    const found = voices.find(
      (voice) =>
        voice.lang.toLowerCase().startsWith("en") &&
        voice.name.toLowerCase().includes(hint.toLowerCase()),
    );
    if (found) {
      preferredVoice = found;
      return found;
    }
  }

  const fallback =
    voices.find((voice) => voice.default && voice.lang.toLowerCase().startsWith("en")) ||
    voices.find((voice) => voice.lang.toLowerCase().startsWith("en")) ||
    voices[0];
  preferredVoice = fallback || null;
  return preferredVoice;
}

function getVoiceQuality() {
  const raw = String(voiceQualityEl?.value || state.voiceQuality || "medium")
    .trim()
    .toLowerCase();
  if (raw === "low" || raw === "medium" || raw === "high") {
    return raw;
  }
  return "medium";
}

function getSpeechMode() {
  return getVoiceQuality() === "high" ? "high" : "browser";
}

function getSpeechVoice() {
  const voice = String(speechVoiceEl?.value || state.voiceName || "shimmer")
    .trim()
    .toLowerCase();
  return /^[a-z0-9_-]{2,40}$/.test(voice) ? voice : "shimmer";
}

function shouldAutoDowngradeHighVoice(errorMessage) {
  const text = String(errorMessage || "").toLowerCase();
  if (!text) return false;
  return (
    text.includes("openai_api_key") ||
    text.includes("incorrect api key") ||
    text.includes("invalid_api_key") ||
    text.includes("quota") ||
    text.includes("billing") ||
    text.includes("model") ||
    text.includes("http 401") ||
    text.includes("http 403") ||
    text.includes("http 429")
  );
}

function getSessionPace() {
  const raw = String(sessionPaceEl?.value || state.sessionPace || "standard")
    .trim()
    .toLowerCase();
  return raw === "busy" ? "busy" : "standard";
}

function getAutoSendVoice() {
  return Boolean(autoSendVoiceEl?.checked);
}

async function unlockAudioPlayback() {
  if (audioPlaybackUnlocked) return true;
  const Ctx = window.AudioContext || window.webkitAudioContext;
  if (!Ctx) {
    audioPlaybackUnlocked = true;
    return true;
  }

  try {
    if (!ttsAudioContext) {
      ttsAudioContext = new Ctx();
    }
    if (ttsAudioContext.state === "suspended") {
      await ttsAudioContext.resume();
    }
    const oscillator = ttsAudioContext.createOscillator();
    const gain = ttsAudioContext.createGain();
    gain.gain.value = 0;
    oscillator.connect(gain);
    gain.connect(ttsAudioContext.destination);
    oscillator.start();
    oscillator.stop(ttsAudioContext.currentTime + 0.01);
    audioPlaybackUnlocked = true;
    return true;
  } catch {
    return false;
  }
}

function installAudioUnlockOnGesture() {
  const unlock = () => {
    void unlockAudioPlayback();
  };
  window.addEventListener("pointerdown", unlock, { passive: true });
  window.addEventListener("keydown", unlock, { passive: true });
  window.addEventListener("touchstart", unlock, { passive: true });
}

function persistDraft() {
  try {
    localStorage.setItem(DRAFT_STORAGE_KEY, String(inputEl?.value || ""));
  } catch {
    // localStorage may be unavailable in privacy mode.
  }
}

function clearDraft() {
  try {
    localStorage.removeItem(DRAFT_STORAGE_KEY);
  } catch {
    // localStorage may be unavailable in privacy mode.
  }
}

function loadDraft() {
  try {
    const draft = localStorage.getItem(DRAFT_STORAGE_KEY);
    if (typeof draft === "string" && draft) {
      inputEl.value = draft;
    }
  } catch {
    // localStorage may be unavailable in privacy mode.
  }
}

function persistSpeechSettings() {
  try {
    localStorage.setItem("consilium.voiceQuality", getVoiceQuality());
    localStorage.setItem("consilium.voiceName", getSpeechVoice());
    localStorage.setItem("consilium.sessionPace", getSessionPace());
    localStorage.setItem("consilium.autoSendVoice", getAutoSendVoice() ? "1" : "0");
  } catch {
    // localStorage may be unavailable in privacy mode.
  }
}

function loadSpeechSettings() {
  try {
    const quality = localStorage.getItem("consilium.voiceQuality");
    const mode = localStorage.getItem("consilium.voiceMode");
    const voiceName = localStorage.getItem("consilium.voiceName");
    const pace = localStorage.getItem("consilium.sessionPace");
    const autoSendVoice = localStorage.getItem("consilium.autoSendVoice");
    if (voiceQualityEl) {
      let nextQuality = "";
      if (quality) {
        nextQuality = ["low", "medium", "high"].includes(quality.toLowerCase())
          ? quality.toLowerCase()
          : "";
      }
      if (!nextQuality && mode) {
        // Backwards compatibility for previous setting.
        nextQuality = mode.toLowerCase() === "high" ? "high" : "medium";
      }
      if (nextQuality) {
        voiceQualityEl.value = nextQuality;
      }
    }
    if (voiceName && speechVoiceEl) {
      if (
        Array.from(speechVoiceEl.options).some(
          (opt) => opt.value.toLowerCase() === voiceName.toLowerCase(),
        )
      ) {
        speechVoiceEl.value = voiceName.toLowerCase();
      }
    }
    if (pace && sessionPaceEl) {
      sessionPaceEl.value = pace.toLowerCase() === "busy" ? "busy" : "standard";
    }
    if (autoSendVoiceEl && autoSendVoice) {
      autoSendVoiceEl.checked = autoSendVoice !== "0";
    }
  } catch {
    // localStorage may be unavailable in privacy mode.
  }
  state.voiceQuality = getVoiceQuality();
  state.voiceName = getSpeechVoice();
  state.sessionPace = getSessionPace();
  state.autoSendVoice = getAutoSendVoice();
}

async function requestHighFidelityAudioChunk(text, token, voiceOverride = "") {
  if (!text) return null;
  if (!speakAbortController || speakAbortController.signal.aborted) {
    speakAbortController = new AbortController();
  }
  const response = await fetch("/api/voice", {
    method: "POST",
    credentials: "same-origin",
    headers: {
      "Content-Type": "application/json",
    },
    signal: speakAbortController.signal,
    body: JSON.stringify({
      text,
      voice: String(voiceOverride || state.voiceName || "shimmer").toLowerCase(),
      format: "mp3",
      speed: 1.03,
      instructions:
        "Natural, warm, human executive-coach voice with measured pacing and clear articulation.",
    }),
  });
  if (!response.ok) {
    const raw = await response.text();
    let payload = {};
    try {
      payload = raw ? JSON.parse(raw) : {};
    } catch {
      payload = {};
    }
    throw new Error(payload.error || raw || `Voice request failed (${response.status}).`);
  }
  if (token !== speakToken) return null;
  return response.blob();
}

async function playAudioBlob(blob, token) {
  if (!blob || token !== speakToken) return;
  const objectUrl = URL.createObjectURL(blob);
  const audio = new Audio(objectUrl);
  activeAudioEl = audio;
  let cleanupDone = false;
  const cleanup = () => {
    if (cleanupDone) return;
    cleanupDone = true;
    stopLipSync();
    URL.revokeObjectURL(objectUrl);
    if (activeAudioEl === audio) {
      activeAudioEl = null;
    }
  };

  await new Promise((resolve, reject) => {
    audio.onended = () => {
      cleanup();
      resolve();
    };
    audio.onerror = () => {
      cleanup();
      reject(new Error("Audio playback failed."));
    };
    audio.onplay = () => {
      startAudioLipSync(audio);
    };
    void unlockAudioPlayback();
    if (ttsAudioContext && ttsAudioContext.state === "suspended") {
      ttsAudioContext.resume().catch(() => {});
    }
    audio.play().catch((err) => {
      cleanup();
      const rawMessage = String(err?.message || err?.name || "").toLowerCase();
      if (
        rawMessage.includes("notallowederror") ||
        rawMessage.includes("not allowed by the user agent") ||
        rawMessage.includes("user agent")
      ) {
        reject(
          new Error(
            "Audio playback blocked by browser policy. Click once on the page, then try again.",
          ),
        );
        return;
      }
      reject(err);
    });
  });
}

async function speakWithHighFidelity(text, token) {
  const normalized = String(text || "").trim();
  if (!normalized) return;

  const rolePlaySegments = parseRolePlaySegments(normalized);
  if (rolePlaySegments.length) {
    const voiceAssignments = buildOpenAIVoiceAssignments(
      rolePlaySegments,
      getSpeechVoice(),
    );
    for (const segment of rolePlaySegments) {
      if (token !== speakToken) return;
      const segmentVoice =
        voiceAssignments.get(normalizeSpeakerName(segment.speaker)) ||
        getSpeechVoice();
      const chunks =
        segment.text.length <= 2000
          ? [segment.text]
          : splitSpeechChunks(segment.text, 900);
      for (const chunk of chunks) {
        if (token !== speakToken) return;
        const audioBlob = await requestHighFidelityAudioChunk(
          chunk,
          token,
          segmentVoice,
        );
        if (!audioBlob || token !== speakToken) return;
        await playAudioBlob(audioBlob, token);
      }
    }
    return;
  }

  const chunks =
    normalized.length <= 3200 ? [normalized] : splitSpeechChunks(normalized, 1400);
  for (const chunk of chunks) {
    if (token !== speakToken) return;
    const audioBlob = await requestHighFidelityAudioChunk(
      chunk,
      token,
      getSpeechVoice(),
    );
    if (!audioBlob || token !== speakToken) return;
    await playAudioBlob(audioBlob, token);
  }
}

function speakWithBrowser(text, token, quality = "medium") {
  if (!("speechSynthesis" in window)) {
    return Promise.reject(new Error("Browser speech synthesis is unavailable."));
  }
  const coachVoice = pickPreferredVoice();
  const rolePlaySegments = parseRolePlaySegments(text);
  const normalizedQuality =
    quality === "low" || quality === "medium" || quality === "high"
      ? quality
      : "medium";
  const rateMap = {
    low: 1.03,
    medium: 0.95,
    high: 0.9,
  };
  const pitchMap = {
    low: 1.0,
    medium: 1.04,
    high: 1.03,
  };
  const motionMap = {
    low: { min: 0.1, max: 0.42, intervalMs: 124, peak: 0.58 },
    medium: { min: 0.14, max: 0.6, intervalMs: 102, peak: 0.78 },
    high: { min: 0.18, max: 0.74, intervalMs: 86, peak: 0.92 },
  };

  const browserVoiceAssignments = rolePlaySegments.length
    ? buildBrowserVoiceAssignments(rolePlaySegments, coachVoice)
    : new Map();
  const speechSegments = rolePlaySegments.length
    ? rolePlaySegments.flatMap((segment) => {
        const chunks = splitSpeechChunks(segment.text, 180);
        return chunks.map((chunk) => ({
          speaker: segment.speaker,
          text: chunk,
          isRolePlay: true,
        }));
      })
    : splitSpeechChunks(text, 210).map((chunk) => ({
        speaker: "Consilium",
        text: chunk,
        isRolePlay: false,
      }));

  let idx = 0;

  return new Promise((resolve, reject) => {
    const speakNext = () => {
      if (token !== speakToken) {
        resolve();
        return;
      }
      if (idx >= speechSegments.length) {
        stopLipSync();
        resolve();
        return;
      }
      const segment = speechSegments[idx];
      idx += 1;
      const utterance = new SpeechSynthesisUtterance(segment.text);

      const speakerKey = normalizeSpeakerName(segment.speaker);
      const isCoach = isCoachSpeaker(speakerKey);
      const speakerVoice =
        browserVoiceAssignments.get(speakerKey) || coachVoice || null;
      if (speakerVoice) utterance.voice = speakerVoice;

      const baseRate = rateMap[normalizedQuality];
      const basePitch = pitchMap[normalizedQuality];
      if (segment.isRolePlay && !isCoach) {
        utterance.rate = Math.max(0.84, Math.min(1.12, baseRate * 0.98));
        utterance.pitch = Math.max(0.86, Math.min(1.24, basePitch * 1.06));
      } else {
        utterance.rate = baseRate;
        utterance.pitch = basePitch;
      }

      const motion = motionMap[normalizedQuality];
      utterance.onstart = () => {
        startFallbackLipMotion(motion);
        pulseMouth(110, motion.peak);
      };
      utterance.onboundary = (event) => {
        const boundaryType = String(event?.name || "").toLowerCase();
        if (boundaryType === "word" || Number(event?.charLength || 0) > 0) {
          pulseMouth(95, motion.peak);
        }
      };
      utterance.onerror = () => {
        stopLipSync();
        reject(new Error("Browser speech playback failed."));
      };
      utterance.onend = () => {
        speakNext();
      };
      window.speechSynthesis.speak(utterance);
    };
    speakNext();
  });
}

async function speak(text) {
  const spoken = normalizeSpeechText(text);
  if (!spoken) return;

  stopSpeaking();
  const token = ++speakToken;
  state.speaking = true;
  refreshCoachVisualState();

  state.voiceQuality = getVoiceQuality();
  state.voiceName = getSpeechVoice();
  persistSpeechSettings();
  const speechMode = getSpeechMode();

  try {
    if (speechMode === "high") {
      await speakWithHighFidelity(spoken, token);
    } else {
      await speakWithBrowser(spoken, token, state.voiceQuality);
    }
  } catch (err) {
    if (token === speakToken && speechMode === "high") {
      const detail = String(err?.message || "").trim();
      const shortDetail =
        detail.length > 120 ? `${detail.slice(0, 117)}...` : detail;
      if (shouldAutoDowngradeHighVoice(detail)) {
        state.voiceQuality = "medium";
        if (voiceQualityEl) voiceQualityEl.value = "medium";
        persistSpeechSettings();
      }
      setStatus(
        shortDetail
          ? `High-fidelity voice unavailable (${shortDetail}). Using browser voice.`
          : "High-fidelity voice unavailable, using browser voice.",
      );
      try {
        await speakWithBrowser(spoken, token, "medium");
      } catch (fallbackErr) {
        if (token === speakToken) {
          setStatus(`Voice error: ${fallbackErr.message}`);
        }
      }
    } else if (token === speakToken) {
      setStatus(`Voice error: ${err.message}`);
    }
  } finally {
    if (token === speakToken) {
      stopLipSync();
      stopAudioPlayback();
      state.speaking = false;
      refreshCoachVisualState();
    }
  }
}

function stopSpeaking() {
  speakToken += 1;
  if ("speechSynthesis" in window) {
    window.speechSynthesis.cancel();
  }
  stopAudioPlayback();
  stopLipSync();
  state.speaking = false;
  refreshCoachVisualState();
}

function clearMessages() {
  state.messages = [];
  messagesEl.innerHTML = "";
}

function resetChat({ preserveDraft = false } = {}) {
  clearMessages();
  stopSpeaking();
  clearFeedbackInputs();
  setFeedbackStatus("");
  if (preserveDraft) {
    loadDraft();
  } else {
    inputEl.value = STARTER_PROMPT;
    clearDraft();
  }
  if (meetingNotesInputEl) meetingNotesInputEl.value = "";
  if (selfReportInputEl) selfReportInputEl.value = "";
  syncAccessState();
  if (canUseCoach()) {
    renderMessage("assistant", OPENING_MESSAGE, true);
  }
  if (preserveDraft && inputEl.value.trim()) {
    setStatus("Draft restored");
  }
  refreshCoachVisualState();
  inputEl.focus();
}

function showSplashPage() {
  mainHeader.classList.add("hidden");
  splashPage.classList.remove("hidden");
  authPanel.classList.add("hidden");
  accountPanel.classList.add("hidden");
  dashboardPage.classList.add("hidden");
  adminPage.classList.add("hidden");
  chatPage.classList.add("hidden");
  dashboardMenuBtn.classList.add("hidden");
  adminMenuBtn.classList.add("hidden");
  accountMenuBtn.classList.add("hidden");
  setCoachVisualState("idle", "Ready to coach");
}

function showAuthPage() {
  mainHeader.classList.remove("hidden");
  splashPage.classList.add("hidden");
  authPanel.classList.remove("hidden");
  accountPanel.classList.add("hidden");
  dashboardPage.classList.add("hidden");
  adminPage.classList.add("hidden");
  chatPage.classList.add("hidden");
  dashboardMenuBtn.classList.add("hidden");
  adminMenuBtn.classList.add("hidden");
  accountMenuBtn.classList.add("hidden");
  setCoachVisualState("idle", "Ready to coach");
}

function applyUserToUI(user) {
  state.user = user;
  accountNameEl.textContent = user.displayName || "Account";
  accountUsernameEl.textContent = user.username ? `@${user.username}` : "";
  accountEmailEl.textContent = user.email || "";
  profileNameEl.value = user.displayName || "";
  profileRoleEl.value = user.role || "";
  profileFocusEl.value = user.focusAreas || "";
  if (supervisorModeToggleEl) {
    const supervisorEnabled = Boolean(state.appSettings?.supervisorModeEnabled);
    supervisorModeToggleEl.checked = Boolean(user.supervisorMode && supervisorEnabled);
    supervisorModeToggleEl.disabled = !supervisorEnabled;
  }
  coachingToneEl.value = normalizeCoachingTone(user.coachingTone);
  if (sessionPaceEl) {
    sessionPaceEl.value = getSessionPace();
  }
  if (adminMenuBtn) {
    adminMenuBtn.classList.toggle("hidden", !Boolean(user.isAdmin));
  }
  applyCoachStyleWeightsToInputs(user.coachStyleWeights);
  applyPersonaAndPlaybookSelectors(user);
  const nudgeSummary = user.nudgeSummary || {};
  if (nudgeStatusEl && state.user) {
    const due = Number(nudgeSummary.due || 0);
    if (due > 0) {
      setNudgeStatus(`${due} nudge${due === 1 ? "" : "s"} due.`);
    } else if (Number(nudgeSummary.total || 0) > 0) {
      setNudgeStatus(`${nudgeSummary.total} nudge${nudgeSummary.total === 1 ? "" : "s"} scheduled.`);
    }
  }
  const practiceSummary = user.practicePlansSummary || {};
  if (practicePlanStatusEl) {
    const total = Number(practiceSummary.total || 0);
    if (total > 0) {
      setPracticePlanStatus(
        `${practiceSummary.active || 0} active, ${practiceSummary.completed || 0} completed (${total} total).`,
      );
    } else {
      setPracticePlanStatus("No practice plans yet.");
    }
  }
  setPasswordStatus(
    user.mustChangePassword ? "Password change required now." : "",
  );
  syncAccessState();
}

function showAuthMode(mode) {
  const registrationDisabled = Boolean(state.appSettings?.registrationDisabled);
  if (mode === "login") {
    loginForm.classList.remove("hidden");
    registerForm.classList.add("hidden");
    showLoginBtn.classList.add("primary");
    showRegisterBtn.classList.remove("primary");
    return;
  }
  if (registrationDisabled) {
    loginForm.classList.remove("hidden");
    registerForm.classList.add("hidden");
    showLoginBtn.classList.add("primary");
    showRegisterBtn.classList.remove("primary");
    setAuthStatus("New account creation is currently disabled.");
    return;
  }
  loginForm.classList.add("hidden");
  registerForm.classList.remove("hidden");
  showLoginBtn.classList.remove("primary");
  showRegisterBtn.classList.add("primary");
}

function showAuthenticatedUI(user, settings = null) {
  mainHeader.classList.remove("hidden");
  splashPage.classList.add("hidden");
  authPanel.classList.add("hidden");
  dashboardMenuBtn.classList.remove("hidden");
  adminMenuBtn.classList.toggle("hidden", !Boolean(user.isAdmin));
  accountMenuBtn.classList.remove("hidden");
  setAuthStatus("");
  setProfileStatus("Profile loaded.");
  if (settings) {
    applyAppSettings(settings);
  }
  applyUserToUI(user);
  resetChat({ preserveDraft: true });
  clearFeedbackInputs();
  setFeedbackStatus("");
  clearNudgeInputs();
  clearPracticePlanInputs();
  setAdminCreateUserStatus("");
  loadCoachLibraries({ quiet: true });
  loadNudges({ quiet: true });
  loadPracticePlans({ quiet: true });
  if (user.mustChangePassword) {
    setView("settings");
  } else {
    setView("chat");
  }
}

async function runQuickPrompt(prompt) {
  if (!prompt) return;
  if (!state.user) {
    setStatus("Sign in first");
    return;
  }
  if (!canUseCoach()) {
    setStatus("Change password to unlock coaching");
    return;
  }
  if (sendBtn.disabled || coachStartBtn.disabled) return;
  inputEl.value = String(prompt).trim();
  persistDraft();
  await sendMessage();
}

function showSignedOutUI({ showSplash = true } = {}) {
  state.user = null;
  state.view = "chat";
  state.nudges = [];
  pendingVoiceAutoSend = false;
  stopSpeaking();
  state.listening = false;
  micStopRequested = false;
  speechRecognitionBaseText = "";
  speechRecognitionFinalText = "";
  if (micBtn) {
    micBtn.textContent = "Start Mic";
  }
  clearMessages();
  inputEl.value = "";
  if (meetingNotesInputEl) meetingNotesInputEl.value = "";
  if (selfReportInputEl) selfReportInputEl.value = "";
  setAuthStatus("Sign in to start coaching.");
  setProfileStatus("");
  setPasswordStatus("");
  setNudgeStatus("");
  setPracticePlanStatus("");
  setFeedbackStatus("");
  setAdminSettingsStatus("");
  setAdminCreateUserStatus("");
  clearFeedbackInputs();
  if (nudgeListEl) nudgeListEl.innerHTML = "";
  if (practicePlanListEl) practicePlanListEl.innerHTML = "";
  if (personaLibraryListEl) personaLibraryListEl.innerHTML = "";
  if (playbookLibraryListEl) playbookLibraryListEl.innerHTML = "";
  state.personas = [];
  state.playbooks = [];
  state.practicePlans = [];
  coachingToneEl.value = DEFAULT_COACHING_TONE;
  if (supervisorModeToggleEl) {
    supervisorModeToggleEl.checked = false;
    supervisorModeToggleEl.disabled = !Boolean(state.appSettings?.supervisorModeEnabled);
  }
  if (stakeholderPersonaSelectEl) stakeholderPersonaSelectEl.value = "";
  if (domainPlaybookSelectEl) domainPlaybookSelectEl.value = "";
  if (practicePlanPersonaSelectEl) practicePlanPersonaSelectEl.value = "";
  if (practicePlanDomainSelectEl) practicePlanDomainSelectEl.value = "";
  if (sessionPaceEl) sessionPaceEl.value = getSessionPace();
  state.sessionPace = getSessionPace();
  applyCoachStyleWeightsToInputs(DEFAULT_STYLE_WEIGHTS);
  if (showSplash) {
    showSplashPage();
  } else {
    showAuthPage();
  }
  adminPage.classList.add("hidden");
  adminMenuBtn.classList.add("hidden");
  syncAccessState();
  setCoachVisualState("idle", "Sign in to begin");
  applyAppSettings(state.appSettings);
}

async function apiRequest(url, options = {}) {
  let response;
  try {
    response = await fetch(url, {
      credentials: "same-origin",
      ...options,
      headers: {
        "Content-Type": "application/json",
        ...(options.headers || {}),
      },
    });
  } catch (err) {
    throw new Error(
      `Network error: ${err.message}. Is the local server running on ${window.location.origin}?`,
    );
  }

  const raw = await response.text();
  let payload = {};
  try {
    payload = raw ? JSON.parse(raw) : {};
  } catch {
    payload = {};
  }

  if (!response.ok) {
    const detail = payload.error || raw || "Request failed.";
    throw new Error(`HTTP ${response.status}: ${detail}`);
  }
  return payload;
}

async function sendMessage() {
  if (!state.user) {
    setStatus("Sign in first");
    return;
  }
  if (!canUseCoach()) {
    setStatus("Change password to unlock coaching");
    return;
  }

  const content = inputEl.value.trim();
  if (!content) return;

  const voiceTurnAuto = pendingVoiceAutoSend;
  pendingVoiceAutoSend = false;

  void unlockAudioPlayback();
  stopSpeaking();
  const selectedPace = getSessionPace();
  state.sessionPace =
    voiceTurnAuto && selectedPace === "standard" ? "busy" : selectedPace;
  renderMessage("user", content, true);
  inputEl.value = "";
  clearDraft();
  sendBtn.disabled = true;
  coachStartBtn.disabled = true;
  setStatus("Thinking...");
  setCoachVisualState("idle", "Thinking...");

  try {
    const meetingNotes = voiceTurnAuto
      ? ""
      : meetingNotesInputEl
        ? meetingNotesInputEl.value.trim()
        : "";
    const selfReport = voiceTurnAuto
      ? ""
      : selfReportInputEl
        ? selfReportInputEl.value.trim()
        : "";
    const contextLimit = voiceTurnAuto
      ? VOICE_CONTEXT_MESSAGES
      : MAX_CONTEXT_MESSAGES;
    const payload = await apiRequest("/api/chat", {
      method: "POST",
      body: JSON.stringify({
        messages: state.messages.slice(-contextLimit),
        meetingNotes,
        selfReport,
        sessionPace: state.sessionPace,
        responseMode: voiceTurnAuto ? "voice_fast" : "standard",
      }),
    });

    renderMessage("assistant", payload.reply, true);
    if (autoSpeakEl.checked) speak(payload.reply);
    setStatus("Ready");
  } catch (err) {
    renderMessage("assistant", `Error: ${err.message}`, false);
    if (err.message.includes("Please log in first.")) {
      showSignedOutUI();
    }
    if (err.message.includes("Password change required")) {
      if (state.user) {
        state.user.mustChangePassword = true;
      }
      syncAccessState();
    }
    setStatus("Error");
  } finally {
    sendBtn.disabled = false;
    coachStartBtn.disabled = false;
    refreshCoachVisualState();
    inputEl.focus();
  }
}

async function initiateCoachConversation() {
  if (!state.user) {
    setStatus("Sign in first");
    return;
  }
  if (!canUseCoach()) {
    setStatus("Change password to unlock coaching");
    return;
  }

  sendBtn.disabled = true;
  coachStartBtn.disabled = true;
  setStatus("Coach is opening...");
  setCoachVisualState("idle", "Opening...");

  try {
    void unlockAudioPlayback();
    stopSpeaking();
    state.sessionPace = getSessionPace();
    const meetingNotes = meetingNotesInputEl ? meetingNotesInputEl.value.trim() : "";
    const selfReport = selfReportInputEl ? selfReportInputEl.value.trim() : "";
    const payload = await apiRequest("/api/chat", {
      method: "POST",
      body: JSON.stringify({
        messages: state.messages.slice(-MAX_CONTEXT_MESSAGES),
        meetingNotes,
        selfReport,
        initiate: true,
        sessionPace: state.sessionPace,
      }),
    });

    renderMessage("assistant", payload.reply, true);
    if (autoSpeakEl.checked) speak(payload.reply);
    setStatus("Ready");
  } catch (err) {
    renderMessage("assistant", `Error: ${err.message}`, false);
    setStatus("Error");
  } finally {
    sendBtn.disabled = false;
    coachStartBtn.disabled = false;
    refreshCoachVisualState();
    inputEl.focus();
  }
}

function mapSpeechRecognitionError(code) {
  const raw = String(code || "").trim().toLowerCase();
  if (!raw) return "Mic error";
  if (raw === "not-allowed" || raw === "service-not-allowed") {
    return "Microphone permission blocked. Allow microphone access in browser/site settings.";
  }
  if (raw === "no-speech") {
    return "No speech detected. Try again and speak clearly after pressing Start Mic.";
  }
  if (raw === "audio-capture") {
    return "No microphone input detected. Check your mic device selection.";
  }
  if (raw === "network") {
    return "Speech recognition network issue. Check your connection and try again.";
  }
  if (raw === "aborted") {
    return "Mic capture stopped.";
  }
  return `Mic error: ${raw}`;
}

async function ensureMicrophonePermission() {
  if (micPermissionVerified) return { ok: true };
  const mediaDevices = navigator.mediaDevices;
  if (!mediaDevices || typeof mediaDevices.getUserMedia !== "function") {
    return { ok: true };
  }

  try {
    if (navigator.permissions && typeof navigator.permissions.query === "function") {
      const result = await navigator.permissions
        .query({ name: "microphone" })
        .catch(() => null);
      if (result && result.state === "denied") {
        return {
          ok: false,
          message:
            "Microphone permission is denied in browser settings. Please allow mic access for this site.",
        };
      }
    }

    const stream = await mediaDevices.getUserMedia({ audio: true });
    for (const track of stream.getTracks()) {
      track.stop();
    }
    micPermissionVerified = true;
    return { ok: true };
  } catch (err) {
    const name = String(err?.name || "").trim();
    if (name === "NotAllowedError" || name === "SecurityError") {
      return {
        ok: false,
        message:
          "Microphone permission blocked. Please allow microphone access, then try again.",
      };
    }
    if (name === "NotFoundError") {
      return {
        ok: false,
        message: "No microphone found. Connect a microphone and try again.",
      };
    }
    if (name === "NotReadableError") {
      return {
        ok: false,
        message: "Microphone is in use by another app. Close other apps using the mic and retry.",
      };
    }
    return {
      ok: false,
      message: `Microphone setup failed: ${err?.message || "Unknown error"}`,
    };
  }
}

function initSpeechRecognition() {
  const SpeechRecognition =
    window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    micBtn.disabled = true;
    micBtn.title = "Speech recognition is not supported in this browser.";
    setStatus("Mic unavailable in this browser. Use Chrome/Edge/Safari.");
    return;
  }

  const recognition = new SpeechRecognition();
  recognition.lang = "en-US";
  recognition.continuous = false;
  recognition.interimResults = true;
  recognition.maxAlternatives = 1;

  recognition.onstart = () => {
    state.listening = true;
    micStopRequested = false;
    speechRecognitionBaseText = String(inputEl?.value || "").trim();
    speechRecognitionFinalText = "";
    micBtn.textContent = "Stop Mic";
    setStatus("Listening... speak now.");
    refreshCoachVisualState();
  };

  recognition.onresult = (event) => {
    let interim = "";
    for (let i = event.resultIndex; i < event.results.length; i += 1) {
      const result = event.results[i];
      const transcript = String(result?.[0]?.transcript || "")
        .replace(/\s+/g, " ")
        .trim();
      if (!transcript) continue;
      if (result.isFinal) {
        speechRecognitionFinalText = speechRecognitionFinalText
          ? `${speechRecognitionFinalText} ${transcript}`
          : transcript;
      } else {
        interim = interim ? `${interim} ${transcript}` : transcript;
      }
    }

    const pieces = [
      speechRecognitionBaseText,
      speechRecognitionFinalText,
      interim,
    ].filter(Boolean);
    inputEl.value = pieces.join(" ").replace(/\s+/g, " ").trim();
    persistDraft();
    inputEl.focus();
  };

  recognition.onend = () => {
    state.listening = false;
    micBtn.textContent = "Start Mic";
    const transcript = String(inputEl?.value || "").trim();
    const hasTranscript =
      transcript.length >= AUTO_SEND_MIN_CHARS || Boolean(speechRecognitionFinalText);
    const autoSendEnabled = getAutoSendVoice();
    const shouldAutoSend =
      autoSendEnabled &&
      hasTranscript &&
      canUseCoach() &&
      !sendBtn.disabled &&
      !coachStartBtn.disabled;

    if (shouldAutoSend) {
      setStatus("Heard you. Thinking...");
    } else if (!hasTranscript) {
      setStatus(
        "Mic stopped with no transcript. Check microphone permission and speak after pressing Start Mic.",
      );
    } else if (!autoSendEnabled && hasTranscript && canUseCoach()) {
      setStatus("Transcript ready. Press Enter to send.");
    } else if (canUseCoach()) {
      setStatus("Ready");
    }
    speechRecognitionBaseText = "";
    speechRecognitionFinalText = "";
    micStopRequested = false;
    refreshCoachVisualState();
    if (shouldAutoSend) {
      pendingVoiceAutoSend = true;
      void sendMessage();
    }
  };

  recognition.onerror = (event) => {
    state.listening = false;
    micBtn.textContent = "Start Mic";
    const errorCode = String(event?.error || "");
    if (
      errorCode === "not-allowed" ||
      errorCode === "service-not-allowed"
    ) {
      micPermissionVerified = false;
    }
    setStatus(mapSpeechRecognitionError(errorCode));
    refreshCoachVisualState();
  };

  recognition.onnomatch = () => {
    setStatus("Could not recognize speech. Try again with clearer audio.");
  };

  recognition.onspeechend = () => {
    if (state.listening && !micStopRequested) {
      setStatus("Got it. Processing...");
      try {
        recognition.stop();
      } catch {
        // Ignore occasional stop races from browser speech APIs.
      }
    }
  };

  state.recognition = recognition;
}

async function toggleMic() {
  if (!canUseCoach() || !state.recognition) return;
  if (state.listening) {
    micStopRequested = true;
    state.recognition.stop();
    return;
  }

  void unlockAudioPlayback();
  setStatus("Checking microphone...");
  const permission = await ensureMicrophonePermission();
  if (!permission.ok) {
    setStatus(permission.message || "Microphone permission check failed.");
    return;
  }

  try {
    state.recognition.start();
  } catch (err) {
    setStatus(`Mic start failed: ${err?.message || "Unknown error"}`);
  }
}

async function checkSession() {
  try {
    const payload = await apiRequest("/api/me", { method: "GET" });
    if (payload.settings) {
      applyAppSettings(payload.settings);
    }
    if (payload.authenticated && payload.user) {
      showAuthenticatedUI(payload.user, payload.settings);
      return;
    }
    showSignedOutUI({ showSplash: true });
  } catch {
    showSignedOutUI({ showSplash: true });
  }
}

async function handleLogin(event) {
  event.preventDefault();
  setAuthStatus("Signing in...");

  const identifier = document.getElementById("loginIdentifier").value.trim();
  const password = document.getElementById("loginPassword").value;

  try {
    const payload = await apiRequest("/api/login", {
      method: "POST",
      body: JSON.stringify({ identifier, password }),
    });
    if (payload.settings) {
      applyAppSettings(payload.settings);
    }
    showAuthenticatedUI(payload.user, payload.settings);
  } catch (err) {
    setAuthStatus(err.message);
  }
}

async function handleRegister(event) {
  event.preventDefault();
  setAuthStatus("Creating account...");

  const username = document.getElementById("registerUsername").value.trim();
  const email = document.getElementById("registerEmail").value.trim();
  const password = document.getElementById("registerPassword").value;
  const displayName = document.getElementById("registerName").value.trim();
  const role = document.getElementById("registerRole").value.trim();
  const focusAreas = document.getElementById("registerFocus").value.trim();

  try {
    const payload = await apiRequest("/api/register", {
      method: "POST",
      body: JSON.stringify({
        username,
        email,
        password,
        displayName,
        role,
        focusAreas,
        coachingTone: normalizeCoachingTone(coachingToneEl.value),
        coachStyleWeights: getCoachStyleWeightsFromInputs(),
      }),
    });
    if (payload.settings) {
      applyAppSettings(payload.settings);
    }
    showAuthenticatedUI(payload.user, payload.settings);
  } catch (err) {
    setAuthStatus(err.message);
  }
}

async function handleLogout() {
  try {
    await apiRequest("/api/logout", { method: "POST", body: "{}" });
  } finally {
    showSignedOutUI({ showSplash: true });
  }
}

async function handleProfileSave() {
  if (!state.user) return;
  setProfileStatus("Saving...");

  try {
    const payload = await apiRequest("/api/profile", {
      method: "PUT",
      body: JSON.stringify({
        displayName: profileNameEl.value.trim(),
        role: profileRoleEl.value.trim(),
        focusAreas: profileFocusEl.value.trim(),
        supervisorMode: Boolean(supervisorModeToggleEl?.checked),
        selectedStakeholderPersonaId: String(
          stakeholderPersonaSelectEl?.value || "",
        ),
        selectedPlaybookId: String(domainPlaybookSelectEl?.value || ""),
        coachingTone: normalizeCoachingTone(coachingToneEl.value),
        coachStyleWeights: getCoachStyleWeightsFromInputs(),
      }),
    });
    applyUserToUI(payload.user);
    setProfileStatus("Profile updated.");
  } catch (err) {
    setProfileStatus(err.message);
  }
}

async function handlePasswordChange() {
  if (!state.user) return;
  const currentPassword = currentPasswordEl.value;
  const newPassword = newPasswordEl.value;
  const confirmPassword = confirmPasswordEl.value;

  if (!newPassword || newPassword.length < 10) {
    setPasswordStatus("New password must be at least 10 characters.");
    return;
  }
  if (newPassword !== confirmPassword) {
    setPasswordStatus("New password and confirm password do not match.");
    return;
  }

  setPasswordStatus("Updating password...");
  try {
    const payload = await apiRequest("/api/password", {
      method: "PUT",
      body: JSON.stringify({ currentPassword, newPassword }),
    });
    applyUserToUI(payload.user);
    currentPasswordEl.value = "";
    newPasswordEl.value = "";
    confirmPasswordEl.value = "";
    setPasswordStatus("Password updated.");
    if (!payload.user.mustChangePassword) {
      setView("chat");
    }
  } catch (err) {
    setPasswordStatus(err.message);
  }
}

sendBtn.addEventListener("click", sendMessage);
coachStartBtn.addEventListener("click", initiateCoachConversation);
micBtn.addEventListener("click", toggleMic);
stopSpeakBtn.addEventListener("click", stopSpeaking);
resetBtn.addEventListener("click", resetChat);
getStartedBtn.addEventListener("click", () => {
  showAuthPage();
  showAuthMode("login");
});
showLoginBtn.addEventListener("click", () => showAuthMode("login"));
showRegisterBtn.addEventListener("click", () => showAuthMode("register"));
loginForm.addEventListener("submit", handleLogin);
registerForm.addEventListener("submit", handleRegister);
logoutBtn.addEventListener("click", handleLogout);
saveProfileBtn.addEventListener("click", handleProfileSave);
changePasswordBtn.addEventListener("click", handlePasswordChange);
if (addNudgeBtn) {
  addNudgeBtn.addEventListener("click", createNudge);
}
if (refreshNudgesBtn) {
  refreshNudgesBtn.addEventListener("click", () => loadNudges());
}
if (createPracticePlanBtn) {
  createPracticePlanBtn.addEventListener("click", createPracticePlan);
}
if (refreshPracticePlansBtn) {
  refreshPracticePlansBtn.addEventListener("click", () =>
    loadPracticePlans({ quiet: false }),
  );
}
if (stakeholderPersonaSelectEl) {
  stakeholderPersonaSelectEl.addEventListener("change", () => {
    if (practicePlanPersonaSelectEl && !practicePlanPersonaSelectEl.value) {
      practicePlanPersonaSelectEl.value = stakeholderPersonaSelectEl.value || "";
    }
  });
}
if (domainPlaybookSelectEl) {
  domainPlaybookSelectEl.addEventListener("change", () => {
    if (practicePlanDomainSelectEl && !practicePlanDomainSelectEl.value) {
      practicePlanDomainSelectEl.value = domainPlaybookSelectEl.value || "";
    }
  });
}
if (submitFeedbackBtn) {
  submitFeedbackBtn.addEventListener("click", submitSessionFeedback);
}
if (saveAdminSettingsBtn) {
  saveAdminSettingsBtn.addEventListener("click", saveAdminSettings);
}
if (adminCreateUserBtn) {
  adminCreateUserBtn.addEventListener("click", createAdminUser);
}
accountMenuBtn.addEventListener("click", () => {
  if (!state.user) return;
  setView("settings");
  loadCoachLibraries({ quiet: true });
  loadNudges({ quiet: true });
  loadPracticePlans({ quiet: true });
});
adminMenuBtn.addEventListener("click", () => {
  if (!state.user || !state.user.isAdmin) return;
  showAdminDashboard();
});
dashboardMenuBtn.addEventListener("click", () => {
  if (!state.user) return;
  showDashboard();
});
backToChatBtn.addEventListener("click", () => {
  if (!state.user) return;
  setView("chat");
});
backToChatFromDashboardBtn.addEventListener("click", () => {
  if (!state.user) return;
  setView("chat");
});
backToChatFromAdminBtn.addEventListener("click", () => {
  if (!state.user) return;
  setView("chat");
});
if (voiceQualityEl) {
  voiceQualityEl.addEventListener("change", () => {
    state.voiceQuality = getVoiceQuality();
    persistSpeechSettings();
  });
}
if (speechVoiceEl) {
  speechVoiceEl.addEventListener("change", () => {
    state.voiceName = getSpeechVoice();
    persistSpeechSettings();
  });
}
if (sessionPaceEl) {
  sessionPaceEl.addEventListener("change", () => {
    state.sessionPace = getSessionPace();
    persistSpeechSettings();
  });
}
if (autoSendVoiceEl) {
  autoSendVoiceEl.addEventListener("change", () => {
    state.autoSendVoice = getAutoSendVoice();
    persistSpeechSettings();
  });
}
for (const btn of quickPromptBtns) {
  btn.addEventListener("click", () => {
    runQuickPrompt(btn.dataset.quickPrompt || "");
  });
}

inputEl.addEventListener("keydown", (event) => {
  const isEnter = event.key === "Enter";
  const isShortcut = (event.metaKey || event.ctrlKey) && isEnter;
  const isPlainEnter = isEnter && !event.shiftKey && !event.altKey && !event.ctrlKey && !event.metaKey;
  if ((isShortcut || isPlainEnter) && !event.isComposing) {
    event.preventDefault();
    sendMessage();
  }
});
inputEl.addEventListener("input", () => {
  persistDraft();
});

if ("speechSynthesis" in window) {
  pickPreferredVoice();
  if ("onvoiceschanged" in window.speechSynthesis) {
    window.speechSynthesis.onvoiceschanged = () => {
      preferredVoice = null;
      pickPreferredVoice();
    };
  }
}

loadSpeechSettings();
loadDraft();
installAudioUnlockOnGesture();
clearNudgeInputs();
clearPracticePlanInputs();
clearFeedbackInputs();
clearAdminCreateUserForm();
applyAppSettings(state.appSettings);
initSpeechRecognition();
showAuthMode("login");
checkSession();
