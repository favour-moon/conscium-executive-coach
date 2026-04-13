import http from "http";
import path from "path";
import crypto from "crypto";
import { readFileSync } from "fs";
import { mkdir, readFile, writeFile } from "fs/promises";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function loadLocalEnvFile() {
  const envPath = path.join(__dirname, ".env");
  let raw = "";
  try {
    raw = readFileSync(envPath, "utf8");
  } catch {
    return;
  }

  const lines = raw.split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const idx = trimmed.indexOf("=");
    if (idx <= 0) continue;

    const key = trimmed.slice(0, idx).trim();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) continue;
    if (process.env[key]) continue;

    let value = trimmed.slice(idx + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    process.env[key] = value;
  }
}

loadLocalEnvFile();

const HOST = (process.env.HOST || "0.0.0.0").trim();
const PORT = Number(process.env.PORT || 8787);
const MODEL = process.env.OPENAI_MODEL || "gpt-4o-mini";
const TTS_MODEL = process.env.OPENAI_TTS_MODEL || "gpt-4o-mini-tts";
const MAX_CHAT_HISTORY_MESSAGES = 16;
const DEFAULT_TTS_VOICE = (process.env.OPENAI_TTS_VOICE || "shimmer").trim();
const DEFAULT_TTS_FORMAT = (
  process.env.OPENAI_TTS_FORMAT || "mp3"
).trim().toLowerCase();
const SUPPORTED_TTS_FORMATS = new Set(["mp3", "wav", "opus", "aac", "flac", "pcm"]);
const SESSION_COOKIE = "coach_session";
const SESSION_TTL_MS = 7 * 24 * 60 * 60 * 1000;
const DATA_DIR = path.join(__dirname, "data");
const USERS_FILE = path.join(DATA_DIR, "users.json");
const DEFAULT_ADMIN_USERNAME = "administrator";
const DEFAULT_ADMIN_PASSWORD = "KingCharles3!";
const STYLE_KEYS = ["directive", "collaborative", "facilitative", "passive"];
const COACHING_TONES = ["concise", "challenging", "supportive"];
const DEFAULT_COACHING_TONE = "supportive";
const DEFAULT_COACH_STYLE_WEIGHTS = Object.freeze({
  directive: 2,
  collaborative: 4,
  facilitative: 4,
  passive: 2,
});
const DEFAULT_APP_SETTINGS = Object.freeze({
  registrationDisabled: false,
  supervisorModeEnabled: true,
});
const MAX_PRACTICE_PLANS_PER_USER = 60;
const PRACTICE_PLAN_STATUS = new Set(["active", "paused", "completed"]);
const PRACTICE_PLAN_CADENCE = new Set([
  "daily",
  "weekly",
  "biweekly",
  "monthly",
  "ad_hoc",
]);
const STAKEHOLDER_PERSONAS = Object.freeze([
  {
    id: "skeptical-cfo",
    name: "Skeptical CFO",
    archetype: "Risk and control leader",
    priorities: [
      "Financial risk exposure",
      "Return on investment",
      "Execution feasibility",
    ],
    pressureStyle: "Analytical and demanding",
    guidance:
      "Use quantifiable outcomes, downside scenarios, and explicit trade-off logic.",
  },
  {
    id: "operational-sponsor",
    name: "Operational Sponsor",
    archetype: "Delivery-oriented executive",
    priorities: ["Predictable delivery", "Clear accountability", "Operational continuity"],
    pressureStyle: "Direct and pace-focused",
    guidance:
      "Lead with timelines, ownership, dependencies, and mitigation plans.",
  },
  {
    id: "policy-guardian",
    name: "Policy Guardian",
    archetype: "Compliance and governance authority",
    priorities: ["Regulatory integrity", "Reputation protection", "Decision traceability"],
    pressureStyle: "Cautious and detail-sensitive",
    guidance:
      "Show governance controls, escalation paths, and documented rationale.",
  },
  {
    id: "transformation-advocate",
    name: "Transformation Advocate",
    archetype: "Innovation-minded sponsor",
    priorities: ["Strategic ambition", "Speed of change", "Capability uplift"],
    pressureStyle: "Future-focused and impatient",
    guidance:
      "Frame strategic upside, phased wins, and organizational learning effects.",
  },
  {
    id: "team-morale-defender",
    name: "Team Morale Defender",
    archetype: "People and culture leader",
    priorities: ["Psychological safety", "Fairness and trust", "Sustainable pace"],
    pressureStyle: "Empathic but firm on values",
    guidance:
      "Balance performance accountability with dignity, support, and clear expectations.",
  },
  {
    id: "ministerial-principal",
    name: "Ministerial Principal",
    archetype: "Senior public stakeholder",
    priorities: ["Public impact", "Political credibility", "Delivery confidence"],
    pressureStyle: "High-stakes and outcome-driven",
    guidance:
      "Use concise narrative, confidence intervals, and clear decision asks.",
  },
  {
    id: "cross-functional-peer",
    name: "Cross-Functional Peer",
    archetype: "Influential lateral partner",
    priorities: ["Mutual wins", "Resource fairness", "Interdependency clarity"],
    pressureStyle: "Negotiation-heavy",
    guidance:
      "Clarify shared objectives, boundary conditions, and reciprocal commitments.",
  },
  {
    id: "underperforming-direct-report",
    name: "Underperforming Direct Report",
    archetype: "Capability under pressure",
    priorities: ["Role clarity", "Support and development", "Fair evaluation"],
    pressureStyle: "Defensive or withdrawn",
    guidance:
      "Use specific behavioral evidence, future-focused expectations, and support commitments.",
  },
]);
const DOMAIN_PLAYBOOKS = Object.freeze([
  {
    id: "underperformance-conversations",
    name: "Underperformance Conversations",
    context:
      "Addressing repeated gaps in delivery, behavior, or quality while preserving trust and fairness.",
    coreMoves: [
      "State observable facts and impact",
      "Invite perspective before judgement",
      "Agree explicit standards and timeframes",
      "Define support plus consequences",
    ],
    pitfalls: [
      "Over-softening expectations",
      "Jumping to blame",
      "Leaving actions or timelines vague",
    ],
  },
  {
    id: "stakeholder-alignment",
    name: "Stakeholder Alignment",
    context:
      "Securing alignment across stakeholders with conflicting incentives and incomplete information.",
    coreMoves: [
      "Map interests, influence, and constraints",
      "Surface non-negotiables early",
      "Frame trade-offs explicitly",
      "Land clear next decision owner",
    ],
    pitfalls: ["Assuming shared definitions", "Ignoring latent blockers", "No follow-through cadence"],
  },
  {
    id: "strategic-decision-making",
    name: "Strategic Decision-Making",
    context:
      "Making high-impact choices under uncertainty with speed and defensible reasoning.",
    coreMoves: [
      "Define decision frame and horizon",
      "Identify options and opportunity cost",
      "Assess second-order effects",
      "Commit and set review trigger points",
    ],
    pitfalls: ["Analysis paralysis", "Hidden assumptions", "Weak reversal criteria"],
  },
  {
    id: "delegation-and-empowerment",
    name: "Delegation and Empowerment",
    context: "Scaling leadership impact by assigning ownership without loss of control.",
    coreMoves: [
      "Clarify outcome and decision rights",
      "Agree guardrails and escalation rules",
      "Match support to capability maturity",
      "Review outcomes not activity",
    ],
    pitfalls: ["Micromanagement", "Ambiguous authority", "Delayed feedback loops"],
  },
  {
    id: "crisis-communication",
    name: "Crisis Communication",
    context:
      "Leading communication in volatile conditions where confidence, speed, and clarity matter.",
    coreMoves: [
      "Lead with known facts and unknowns",
      "Set immediate priorities and owners",
      "Create update cadence and channels",
      "Close with confidence and contingency",
    ],
    pitfalls: ["Over-reassurance", "Message inconsistency", "Slow correction of new facts"],
  },
  {
    id: "board-and-exco-briefing",
    name: "Board and ExCo Briefing",
    context:
      "Condensing complexity into strategic narrative and decision-ready recommendations.",
    coreMoves: [
      "Start with decision required",
      "Summarize evidence and risk posture",
      "Present options with trade-offs",
      "State recommendation and ask",
    ],
    pitfalls: ["Too much operational detail", "Unclear recommendation", "No quantified downside"],
  },
]);
const STAKEHOLDER_PERSONA_MAP = new Map(
  STAKEHOLDER_PERSONAS.map((persona) => [persona.id, persona]),
);
const DOMAIN_PLAYBOOK_MAP = new Map(
  DOMAIN_PLAYBOOKS.map((playbook) => [playbook.id, playbook]),
);
const COACHING_CONSTITUTION_VERSION = "2.0";
const ANALYSIS_MODEL = process.env.OPENAI_ANALYSIS_MODEL || MODEL;
const PROFILE_RETENTION_DAYS = 365;
const MAX_NUDGES_PER_USER = 80;
const CATEGORY_KEYS = [
  "decisionSpeed",
  "delegation",
  "conversationQuality",
  "adaptability",
];
const BIG_FIVE_KEYS = [
  "openness",
  "conscientiousness",
  "extraversion",
  "agreeableness",
  "neuroticism",
];
const DEFAULT_CATEGORY_SCORES = Object.freeze({
  decisionSpeed: 5,
  delegation: 5,
  conversationQuality: 5,
  adaptability: 5,
});
const DEFAULT_BIG_FIVE = Object.freeze({
  openness: 5,
  conscientiousness: 5,
  extraversion: 5,
  agreeableness: 5,
  neuroticism: 5,
});
const COACH_STYLE_CONFIG = {
  directive: {
    label: "Directive (Instructing)",
    summary:
      "Provide clear recommendations with explicit reasoning and immediate execution steps.",
    instruction:
      "Lead with a specific recommendation when clarity and speed matter. Minimize ambiguity and make decision rationale explicit.",
  },
  collaborative: {
    label: "Collaborative (Partnership)",
    summary:
      "Act as a thinking partner by pressure-testing assumptions and refining the approach jointly.",
    instruction:
      "Engage in active dialogue that challenges and sharpens the user's viewpoint while preserving ownership.",
  },
  facilitative: {
    label: "Facilitative (Empowering)",
    summary:
      "Use structured questions and reflection to help the leader generate options and build judgement.",
    instruction:
      "Prioritize inquiry and reflection over instruction to strengthen the user's independent thinking.",
  },
  passive: {
    label: "Passive (Non-Directive / Supporting)",
    summary:
      "Create disciplined space with minimal intervention while offering selective reflections.",
    instruction:
      "Intervene lightly, listen carefully, and help the user articulate their own conclusions.",
  },
};
const COACH_TONE_CONFIG = {
  concise:
    "Keep responses brief and crisp, focus on one practical question or suggestion at a time.",
  challenging:
    "Challenge assumptions directly, surface blind spots, and press for clear commitments.",
  supportive:
    "Use warm, encouraging language while still driving toward specific action and accountability.",
};

const SYSTEM_PROMPT = `
You are an executive coach for senior leaders in live operational contexts.
Default interaction style:
- Talk like a human coach in natural dialogue, not rigid templates.
- Keep each turn concise and speakable (typically 2-5 sentences).
- Ask one focused question at a time, then wait for the user's reply.
- Avoid headers, scoring blocks, and long lists unless the user explicitly asks for a structured format.
- Use GROW and Kantor moves internally, but do not force framework labels every turn.
- Do not output stage labels like "Goal / Reality / Options / Will".
- Start with the user's concrete goal for the conversation, then explore reality, options, and commitment through normal dialogue.

Coaching priorities:
- Difficult conversations, decision speed, delegation, and system-level thinking.
- Ground recommendations in the user's real constraints and timelines.
- If urgency is high, be more direct; otherwise use collaborative/facilitative questioning.

Governance rules:
- Explicitly respect consent boundaries.
- Keep raw details private to the individual.
- Never frame coaching as covert monitoring.
- Keep coaching separate from performance-management judgement.
`;

const COACHING_CONSTITUTION_PROMPT = `
Consilium Coaching Constitution (Version ${COACHING_CONSTITUTION_VERSION})

Purpose:
- Improve leadership judgement and behavior in real-world contexts.
- Help leaders think clearly, decide effectively, and act with confidence.

Core philosophy:
- Judgement over answers.
- Action over reflection: every meaningful exchange ends with a decision, behavior change, or explicit next step.
- Reality over abstraction: stay grounded in live context and avoid generic advice.
- Challenge with intent: surface contradictions and weak reasoning proportionately.
- Trust as a foundation: explicit consent, clear data boundaries, user control.

Objectives:
- Decision-making: speed under uncertainty, clarity of reasoning, confidence in action.
- Delegation and empowerment: distribute authority, reduce over-control, build team capability.
- Strategic thinking: system-level awareness, second-order effects, mission alignment.
- Adaptability: learning from feedback and behavior change over time.

Framework integration:
- GROW for session structure.
- Kantor Four Moves for interaction dynamics.
- COM-B for behavior diagnosis.
- Deliberate practice for behavior reinforcement.

Mode adaptation:
- Default to facilitative/collaborative unless context demands otherwise.
- Escalate when urgency or repeated patterns require more direction.
- De-escalate in sensitive or emotional contexts.
- Signal intent when shifting tone.

Behavioral insight and intervention rules:
- Keep feedback evidence-based and acknowledge confidence/visibility limits.
- Prioritize recurring patterns over isolated events.
- Prioritize high-value interventions and limit nudges to 2-3 per week.
- Avoid unnecessary interruption and do not overwhelm the user.

Development loop:
- Observation -> Feedback -> Action -> Application -> Reflection -> Adjustment.
- Keep development actions specific, time-bound, behavior-linked, and limited in number.

Role-play and communication:
- Simulate realistic stakeholders with adaptive responses.
- Always provide feedback on clarity, tone, and effectiveness.
- Communicate concisely, directly, professionally, and without jargon/cliches.

Trust and boundaries:
- Consent governs data use.
- Raw content is private and not exposed.
- Retain only behavioral signals.
- Be transparent about data usage.

Failure modes to avoid:
- Overly passive (no impact), overly directive (loss of trust), generic advice, over-intervention.

Success definition:
- Faster decision-making, increased delegation, improved difficult conversations, and observable behavior change.
`;

const PROFILE_ANALYZER_PROMPT = `
You are the profile and intervention analyst for Consilium.

Task:
- After each coaching interaction, update the coachee profile from evidence.
- Use available evidence only: chat exchange, meeting notes, and self-report.
- Output strict JSON only (no markdown, no commentary).

Requirements:
- Track Big Five (1-10 each), MBTI best-fit, and DISC best-fit.
- Track category scores (1-10): decisionSpeed, delegation, conversationQuality, adaptability.
- Update strengths, developmentAreas, and observable traits.
- Recommend or update interventions and mark status/success.
- Assess whether interventions are working over time.

JSON schema:
{
  "traits": ["..."],
  "strengths": ["..."],
  "developmentAreas": ["..."],
  "personality": {
    "bigFive": {
      "openness": 1-10,
      "conscientiousness": 1-10,
      "extraversion": 1-10,
      "agreeableness": 1-10,
      "neuroticism": 1-10
    },
    "mbti": "string",
    "disc": "string"
  },
  "categoryScores": {
    "decisionSpeed": 1-10,
    "delegation": 1-10,
    "conversationQuality": 1-10,
    "adaptability": 1-10
  },
  "interventions": [
    {
      "id": "optional-existing-id",
      "title": "string",
      "type": "role-play|nudge|practice|reflection|accountability|other",
      "rationale": "string",
      "targetCategories": ["decisionSpeed|delegation|conversationQuality|adaptability"],
      "actions": ["string"],
      "status": "active|improving|succeeded|stalled",
      "successScore": 1-10,
      "evidence": ["string"]
    }
  ],
  "interactionAssessment": {
    "summary": "short summary",
    "progress": "improved|stable|declined|mixed",
    "confidence": "low|medium|high"
  }
}

Rules:
- Be conservative when evidence is weak.
- Prefer recurring patterns over one-off signals.
- Keep list sizes tight and useful.
`;

let userStore = {
  users: [],
  appSettings: { ...DEFAULT_APP_SETTINGS },
};
const sessions = new Map();
let persistQueue = Promise.resolve();

function sendJson(res, code, payload) {
  res.writeHead(code, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(payload));
}

function getContentType(filePath) {
  const ext = path.extname(filePath);
  if (ext === ".html") return "text/html; charset=utf-8";
  if (ext === ".js") return "text/javascript; charset=utf-8";
  if (ext === ".css") return "text/css; charset=utf-8";
  if (ext === ".png") return "image/png";
  if (ext === ".jpg" || ext === ".jpeg") return "image/jpeg";
  if (ext === ".svg") return "image/svg+xml";
  if (ext === ".json") return "application/json; charset=utf-8";
  return "application/octet-stream";
}

async function serveStatic(pathname, res) {
  const publicRoot = path.join(__dirname, "public");
  const reqPath = pathname === "/" ? "/index.html" : pathname;
  const cleanPath = path.normalize(reqPath).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(publicRoot, cleanPath);

  if (!filePath.startsWith(publicRoot)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  try {
    const body = await readFile(filePath);
    res.writeHead(200, { "Content-Type": getContentType(filePath) });
    res.end(body);
  } catch {
    res.writeHead(404);
    res.end("Not found");
  }
}

async function parseBody(req) {
  let data = "";
  for await (const chunk of req) {
    data += chunk;
    if (data.length > 1024 * 1024) {
      throw new Error("Request too large");
    }
  }
  if (!data.trim()) return {};
  return JSON.parse(data);
}

function normalizeEmail(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function normalizeUsername(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function isValidUsername(value) {
  return /^[a-z0-9_-]{3,32}$/.test(value);
}

function clampStyleWeight(value) {
  const parsed = Number.parseInt(String(value), 10);
  if (!Number.isFinite(parsed)) return null;
  return Math.max(1, Math.min(5, parsed));
}

function normalizeCoachingTone(value, fallback = DEFAULT_COACHING_TONE) {
  const candidate = String(value || "")
    .trim()
    .toLowerCase();
  if (COACHING_TONES.includes(candidate)) return candidate;
  return fallback;
}

function getLegacyStyleWeights(style) {
  const key = String(style || "")
    .trim()
    .toLowerCase();
  if (!Object.hasOwn(COACH_STYLE_CONFIG, key)) return null;

  return {
    directive: key === "directive" ? 5 : 2,
    collaborative: key === "collaborative" ? 5 : 2,
    facilitative: key === "facilitative" ? 5 : 2,
    passive: key === "passive" ? 5 : 2,
  };
}

function normalizeCoachStyleWeights(value, legacyStyle = null) {
  const source = value && typeof value === "object" ? value : {};
  const fallback =
    getLegacyStyleWeights(legacyStyle) || DEFAULT_COACH_STYLE_WEIGHTS;

  const normalized = {};
  for (const key of STYLE_KEYS) {
    const fromSource = clampStyleWeight(source[key]);
    normalized[key] = fromSource ?? fallback[key];
  }
  return normalized;
}

function coachStyleWeightsEqual(a, b) {
  return STYLE_KEYS.every((key) => a[key] === b[key]);
}

function getDominantStyles(weights) {
  let max = 0;
  for (const key of STYLE_KEYS) {
    max = Math.max(max, weights[key]);
  }
  return STYLE_KEYS.filter((key) => weights[key] === max);
}

function sanitizeText(value, maxLen = 200) {
  return String(value || "")
    .trim()
    .slice(0, maxLen);
}

function sanitizeArray(value, maxItems = 8, maxLen = 140) {
  if (!Array.isArray(value)) return [];
  const clean = value
    .map((item) => sanitizeText(item, maxLen))
    .filter(Boolean);
  return clean.slice(0, maxItems);
}

function clampScore10(value, fallback = 5) {
  const parsed = Number.parseFloat(value);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(1, Math.min(10, Math.round(parsed)));
}

function normalizeScoreMap(source, keys, defaults) {
  const out = {};
  const input = source && typeof source === "object" ? source : {};
  for (const key of keys) {
    out[key] = clampScore10(input[key], defaults[key]);
  }
  return out;
}

function normalizeDeltaMap(source, keys) {
  const input = source && typeof source === "object" ? source : {};
  const out = {};
  for (const key of keys) {
    const parsed = Number.parseFloat(input[key]);
    if (!Number.isFinite(parsed)) {
      out[key] = 0;
      continue;
    }
    out[key] = Math.max(-10, Math.min(10, Math.round(parsed)));
  }
  return out;
}

function normalizeSessionPace(value) {
  const candidate = sanitizeText(value, 20).toLowerCase();
  return candidate === "busy" ? "busy" : "standard";
}

function inferUseCaseFromText(value, fallback = "General coaching") {
  const text = sanitizeText(value, 400).toLowerCase();
  if (!text) return fallback;
  if (text.includes("underperformance") || text.includes("performance")) {
    return "Underperformance conversation";
  }
  if (text.includes("delegat")) return "Delegation and empowerment";
  if (text.includes("stakeholder")) return "Stakeholder management";
  if (text.includes("role-play") || text.includes("role play")) {
    return "Role-play practice";
  }
  if (text.includes("decision") || text.includes("decide")) return "Decision support";
  if (
    text.includes("feedback") ||
    text.includes("conversation") ||
    text.includes("conflict")
  ) {
    return "Difficult conversation";
  }
  if (text.includes("strategy") || text.includes("strategic")) {
    return "Strategic thinking";
  }
  return fallback;
}

function normalizeTtsVoice(value, fallback = DEFAULT_TTS_VOICE) {
  const candidate = sanitizeText(value, 40).toLowerCase();
  if (!candidate) return fallback;
  if (!/^[a-z0-9_-]+$/.test(candidate)) return fallback;
  return candidate;
}

function normalizeTtsFormat(value, fallback = DEFAULT_TTS_FORMAT) {
  const candidate = sanitizeText(value, 10).toLowerCase();
  if (SUPPORTED_TTS_FORMATS.has(candidate)) return candidate;
  if (SUPPORTED_TTS_FORMATS.has(fallback)) return fallback;
  return "mp3";
}

function getTtsContentType(format) {
  if (format === "wav") return "audio/wav";
  if (format === "opus") return "audio/ogg";
  if (format === "aac") return "audio/aac";
  if (format === "flac") return "audio/flac";
  if (format === "pcm") return "audio/L16";
  return "audio/mpeg";
}

function normalizeAppSettings(value) {
  const source = value && typeof value === "object" ? value : {};
  return {
    registrationDisabled:
      typeof source.registrationDisabled === "boolean"
        ? source.registrationDisabled
        : DEFAULT_APP_SETTINGS.registrationDisabled,
    supervisorModeEnabled:
      typeof source.supervisorModeEnabled === "boolean"
        ? source.supervisorModeEnabled
        : DEFAULT_APP_SETTINGS.supervisorModeEnabled,
  };
}

function getAppSettings() {
  const normalized = normalizeAppSettings(userStore.appSettings);
  userStore.appSettings = normalized;
  return normalized;
}

function normalizeStakeholderPersonaId(value, fallback = "") {
  const candidate = sanitizeText(value, 80).toLowerCase();
  if (candidate && STAKEHOLDER_PERSONA_MAP.has(candidate)) return candidate;
  return fallback;
}

function normalizePlaybookId(value, fallback = "") {
  const candidate = sanitizeText(value, 80).toLowerCase();
  if (candidate && DOMAIN_PLAYBOOK_MAP.has(candidate)) return candidate;
  return fallback;
}

function getStakeholderPersona(personaId) {
  const normalized = normalizeStakeholderPersonaId(personaId, "");
  return normalized ? STAKEHOLDER_PERSONA_MAP.get(normalized) || null : null;
}

function getDomainPlaybook(playbookId) {
  const normalized = normalizePlaybookId(playbookId, "");
  return normalized ? DOMAIN_PLAYBOOK_MAP.get(normalized) || null : null;
}

function normalizePracticePlanStatus(value, fallback = "active") {
  const candidate = sanitizeText(value, 24).toLowerCase();
  return PRACTICE_PLAN_STATUS.has(candidate) ? candidate : fallback;
}

function normalizePracticePlanCadence(value, fallback = "weekly") {
  const candidate = sanitizeText(value, 24).toLowerCase();
  return PRACTICE_PLAN_CADENCE.has(candidate) ? candidate : fallback;
}

function normalizeOptionalIsoDate(value) {
  const ms = toMs(value);
  return ms ? new Date(ms).toISOString() : "";
}

function normalizePracticePlanRecord(item) {
  const now = newIsoNow();
  const completedCountRaw = Number.parseInt(String(item?.completedCount), 10);
  const completedCount =
    Number.isFinite(completedCountRaw) && completedCountRaw > 0
      ? Math.min(1000, completedCountRaw)
      : 0;
  return {
    id: sanitizeText(item?.id, 80) || crypto.randomUUID(),
    title: sanitizeText(item?.title, 140) || "Practice plan",
    objective: sanitizeText(item?.objective, 420),
    domainId: normalizePlaybookId(item?.domainId || item?.playbookId, ""),
    stakeholderPersonaId: normalizeStakeholderPersonaId(
      item?.stakeholderPersonaId || item?.personaId,
      "",
    ),
    actions: sanitizeArray(item?.actions, 8, 200),
    cadence: normalizePracticePlanCadence(item?.cadence, "weekly"),
    status: normalizePracticePlanStatus(item?.status, "active"),
    nextDueAt: normalizeOptionalIsoDate(item?.nextDueAt),
    lastPracticedAt: normalizeOptionalIsoDate(item?.lastPracticedAt),
    completedCount,
    notes: sanitizeText(item?.notes, 420),
    createdAt: sanitizeText(item?.createdAt, 40) || now,
    updatedAt: sanitizeText(item?.updatedAt, 40) || now,
  };
}

function normalizePracticePlans(value) {
  const base = Array.isArray(value) ? value : [];
  const cutoff = retentionCutoffMs();
  return base
    .map((item) => normalizePracticePlanRecord(item))
    .filter((plan) => toMs(plan.updatedAt || plan.createdAt) >= cutoff)
    .sort((a, b) => toMs(b.updatedAt || b.createdAt) - toMs(a.updatedAt || a.createdAt))
    .slice(0, MAX_PRACTICE_PLANS_PER_USER);
}

function getPracticePlanSummary(plans) {
  const list = Array.isArray(plans) ? plans : [];
  const now = Date.now();
  return {
    total: list.length,
    active: list.filter((plan) => plan.status === "active").length,
    paused: list.filter((plan) => plan.status === "paused").length,
    completed: list.filter((plan) => plan.status === "completed").length,
    due: list.filter(
      (plan) => plan.status === "active" && plan.nextDueAt && toMs(plan.nextDueAt) <= now,
    ).length,
  };
}

function normalizeFeedbackLoop(input) {
  const base = input && typeof input === "object" ? input : {};
  const count = (value) => {
    const parsed = Number.parseInt(String(value), 10);
    if (!Number.isFinite(parsed) || parsed < 0) return 0;
    return parsed;
  };
  const source = sanitizeText(base.lastAssessmentSource, 24).toLowerCase();
  return {
    totalInteractions: count(base.totalInteractions),
    analyzedInteractions: count(base.analyzedInteractions),
    fallbackInteractions: count(base.fallbackInteractions),
    lastAssessmentAt: sanitizeText(base.lastAssessmentAt, 40) || "",
    lastAssessmentSource: ["analyzer", "fallback", "none"].includes(source)
      ? source
      : "none",
    lastInterventionReviewAt: sanitizeText(base.lastInterventionReviewAt, 40) || "",
  };
}

function recomputeFeedbackLoop(profile) {
  const loop = normalizeFeedbackLoop(profile.feedbackLoop);
  const total = Array.isArray(profile.interactionLog)
    ? profile.interactionLog.length
    : 0;
  const analyzed = Array.isArray(profile.interactionLog)
    ? profile.interactionLog.filter((entry) => entry.assessmentMode === "analyzer")
        .length
    : 0;
  const fallback = Array.isArray(profile.interactionLog)
    ? profile.interactionLog.filter((entry) => entry.assessmentMode === "fallback")
        .length
    : 0;
  return {
    ...loop,
    totalInteractions: total,
    analyzedInteractions: analyzed,
    fallbackInteractions: fallback,
  };
}

function normalizeOptionalScore10(value) {
  if (value === null || value === undefined || value === "") return null;
  const parsed = Number.parseFloat(value);
  if (!Number.isFinite(parsed)) return null;
  return Math.max(1, Math.min(10, Math.round(parsed)));
}

function normalizeInteractionFeedback(value) {
  const base = value && typeof value === "object" ? value : {};
  const helpfulness = normalizeOptionalScore10(base.helpfulness);
  const clarity = normalizeOptionalScore10(base.clarity);
  const confidence = normalizeOptionalScore10(base.confidence);
  const comment = sanitizeText(base.comment, 400);
  if (
    helpfulness === null &&
    clarity === null &&
    confidence === null &&
    !comment
  ) {
    return null;
  }
  return {
    helpfulness,
    clarity,
    confidence,
    comment,
    submittedAt: sanitizeText(base.submittedAt, 40) || newIsoNow(),
  };
}

function normalizeNudgeStatus(value, fallback = "scheduled") {
  const candidate = sanitizeText(value, 24).toLowerCase();
  if (["scheduled", "due", "completed", "dismissed"].includes(candidate)) {
    return candidate;
  }
  return fallback;
}

function normalizeNudgeRecord(item) {
  const now = newIsoNow();
  const dueAtMs = toMs(item?.dueAt);
  const dueAt = dueAtMs ? new Date(dueAtMs).toISOString() : now;
  return {
    id: sanitizeText(item?.id, 80) || crypto.randomUUID(),
    title: sanitizeText(item?.title, 120) || "Coaching nudge",
    message: sanitizeText(item?.message, 320),
    dueAt,
    status: normalizeNudgeStatus(item?.status, "scheduled"),
    channel: "in_app",
    createdAt: sanitizeText(item?.createdAt, 40) || now,
    updatedAt: sanitizeText(item?.updatedAt, 40) || now,
  };
}

function normalizeNudges(value) {
  const base = Array.isArray(value) ? value : [];
  const cutoff = retentionCutoffMs();
  return base
    .map(normalizeNudgeRecord)
    .filter((item) => toMs(item.updatedAt || item.createdAt) >= cutoff)
    .sort((a, b) => toMs(a.dueAt) - toMs(b.dueAt))
    .slice(0, MAX_NUDGES_PER_USER);
}

function syncUserNudges(user) {
  const nudges = normalizeNudges(user.nudges);
  const now = Date.now();
  let changed = false;
  for (const nudge of nudges) {
    const dueMs = toMs(nudge.dueAt);
    if (nudge.status === "scheduled" && dueMs && dueMs <= now) {
      nudge.status = "due";
      nudge.updatedAt = newIsoNow();
      changed = true;
      continue;
    }
    if (nudge.status === "due" && dueMs && dueMs > now) {
      nudge.status = "scheduled";
      nudge.updatedAt = newIsoNow();
      changed = true;
    }
  }
  user.nudges = nudges;
  return changed;
}

function getNudgeSummary(user) {
  const nudges = Array.isArray(user.nudges) ? user.nudges : [];
  const summary = {
    total: nudges.length,
    scheduled: 0,
    due: 0,
    completed: 0,
    dismissed: 0,
  };
  for (const nudge of nudges) {
    const status = normalizeNudgeStatus(nudge?.status, "scheduled");
    summary[status] += 1;
  }
  return summary;
}

function retentionCutoffMs() {
  return Date.now() - PROFILE_RETENTION_DAYS * 24 * 60 * 60 * 1000;
}

function toMs(value) {
  const ms = Date.parse(String(value || ""));
  return Number.isFinite(ms) ? ms : 0;
}

function newIsoNow() {
  return new Date().toISOString();
}

function hashPassword(password) {
  const salt = crypto.randomBytes(16).toString("hex");
  const hash = crypto.scryptSync(password, salt, 64).toString("hex");
  return `${salt}:${hash}`;
}

function verifyPassword(password, stored) {
  const [salt, storedHash] = String(stored || "").split(":");
  if (!salt || !storedHash) return false;
  const computed = crypto.scryptSync(password, salt, 64).toString("hex");
  const a = Buffer.from(storedHash, "hex");
  const b = Buffer.from(computed, "hex");
  if (a.length !== b.length) return false;
  return crypto.timingSafeEqual(a, b);
}

function createDefaultCoachingProfile() {
  return {
    traits: [],
    strengths: [],
    developmentAreas: [],
    personality: {
      bigFive: { ...DEFAULT_BIG_FIVE },
      mbti: "Unknown",
      disc: "Unknown",
    },
    categoryScores: { ...DEFAULT_CATEGORY_SCORES },
    interventions: [],
    interactionLog: [],
    metrics: {
      averageCategoryScore: 5,
      activeInterventions: 0,
      successfulInterventions: 0,
      interventionSuccessRate: 0,
    },
    feedbackLoop: {
      totalInteractions: 0,
      analyzedInteractions: 0,
      fallbackInteractions: 0,
      lastAssessmentAt: "",
      lastAssessmentSource: "none",
      lastInterventionReviewAt: "",
    },
    updatedAt: newIsoNow(),
  };
}

function normalizeInterventionRecord(item) {
  const now = newIsoNow();
  const status = String(item?.status || "active").toLowerCase();
  const safeStatus = ["active", "improving", "succeeded", "stalled"].includes(status)
    ? status
    : "active";
  return {
    id: sanitizeText(item?.id, 80) || crypto.randomUUID(),
    title: sanitizeText(item?.title, 120) || "Untitled intervention",
    type: sanitizeText(item?.type, 40) || "other",
    rationale: sanitizeText(item?.rationale, 280),
    targetCategories: sanitizeArray(item?.targetCategories, 4, 40).filter((key) =>
      CATEGORY_KEYS.includes(key),
    ),
    actions: sanitizeArray(item?.actions, 6, 180),
    status: safeStatus,
    successScore: clampScore10(item?.successScore, 5),
    evidence: sanitizeArray(item?.evidence, 6, 180),
    createdAt: sanitizeText(item?.createdAt, 40) || now,
    updatedAt: sanitizeText(item?.updatedAt, 40) || now,
  };
}

function computeProfileMetrics(profile) {
  const scores = CATEGORY_KEYS.map((key) => profile.categoryScores[key]);
  const avg =
    scores.reduce((sum, value) => sum + value, 0) / Math.max(1, scores.length);
  const interventions = profile.interventions || [];
  const successful = interventions.filter(
    (item) => item.status === "succeeded" || item.successScore >= 7,
  ).length;
  const active = interventions.filter(
    (item) => item.status === "active" || item.status === "improving",
  ).length;
  const rate = interventions.length
    ? Number(((successful / interventions.length) * 100).toFixed(1))
    : 0;
  return {
    averageCategoryScore: Number(avg.toFixed(1)),
    activeInterventions: active,
    successfulInterventions: successful,
    interventionSuccessRate: rate,
  };
}

function normalizeCoachingProfile(profile) {
  const base =
    profile && typeof profile === "object" ? profile : createDefaultCoachingProfile();
  const normalized = {
    traits: sanitizeArray(base.traits, 12, 80),
    strengths: sanitizeArray(base.strengths, 12, 120),
    developmentAreas: sanitizeArray(base.developmentAreas, 12, 120),
    personality: {
      bigFive: normalizeScoreMap(
        base.personality?.bigFive,
        BIG_FIVE_KEYS,
        DEFAULT_BIG_FIVE,
      ),
      mbti: sanitizeText(base.personality?.mbti || "Unknown", 10) || "Unknown",
      disc: sanitizeText(base.personality?.disc || "Unknown", 10) || "Unknown",
    },
    categoryScores: normalizeScoreMap(
      base.categoryScores,
      CATEGORY_KEYS,
      DEFAULT_CATEGORY_SCORES,
    ),
    interventions: Array.isArray(base.interventions)
      ? base.interventions.map(normalizeInterventionRecord).slice(0, 30)
      : [],
    interactionLog: Array.isArray(base.interactionLog)
      ? base.interactionLog
          .map((entry) => ({
            id: sanitizeText(entry?.id, 80) || crypto.randomUUID(),
            timestamp: sanitizeText(entry?.timestamp, 40) || newIsoNow(),
            summary: sanitizeText(entry?.summary, 320),
            progress: sanitizeText(entry?.progress, 20) || "stable",
            confidence: sanitizeText(entry?.confidence, 20) || "medium",
            assessmentMode:
              sanitizeText(entry?.assessmentMode, 20).toLowerCase() === "fallback"
                ? "fallback"
                : "analyzer",
            sessionPace: normalizeSessionPace(entry?.sessionPace),
            useCase: sanitizeText(entry?.useCase, 120) || "General coaching",
            userPrompt: sanitizeText(entry?.userPrompt, 240),
            coachResponse: sanitizeText(entry?.coachResponse, 240),
            meetingNotesSnippet: sanitizeText(entry?.meetingNotesSnippet, 240),
            selfReportSnippet: sanitizeText(entry?.selfReportSnippet, 240),
            coacheeFeedback: normalizeInteractionFeedback(entry?.coacheeFeedback),
            categoryScores: normalizeScoreMap(
              entry?.categoryScores,
              CATEGORY_KEYS,
              DEFAULT_CATEGORY_SCORES,
            ),
            scoreDelta: normalizeDeltaMap(entry?.scoreDelta, CATEGORY_KEYS),
            evidenceFlags: {
              meetingNotesProvided: Boolean(entry?.evidenceFlags?.meetingNotesProvided),
              selfReportProvided: Boolean(entry?.evidenceFlags?.selfReportProvided),
            },
          }))
          .slice(0, 400)
      : [],
    metrics: base.metrics && typeof base.metrics === "object" ? base.metrics : {},
    feedbackLoop: normalizeFeedbackLoop(base.feedbackLoop),
    updatedAt: sanitizeText(base.updatedAt, 40) || newIsoNow(),
  };

  const cutoff = retentionCutoffMs();
  normalized.interventions = normalized.interventions.filter(
    (item) => toMs(item.updatedAt || item.createdAt) >= cutoff,
  );
  normalized.interactionLog = normalized.interactionLog.filter(
    (entry) => toMs(entry.timestamp) >= cutoff,
  );
  normalized.feedbackLoop = recomputeFeedbackLoop(normalized);
  if (!normalized.feedbackLoop.lastAssessmentAt && normalized.interactionLog.length) {
    normalized.feedbackLoop.lastAssessmentAt =
      normalized.interactionLog[normalized.interactionLog.length - 1].timestamp;
  }
  if (normalized.feedbackLoop.lastAssessmentSource === "none") {
    const lastMode =
      normalized.interactionLog[normalized.interactionLog.length - 1]
        ?.assessmentMode || "";
    if (lastMode === "analyzer" || lastMode === "fallback") {
      normalized.feedbackLoop.lastAssessmentSource = lastMode;
    }
  }
  normalized.metrics = computeProfileMetrics(normalized);
  return normalized;
}

function ensureUserCoachingProfile(user) {
  const normalized = normalizeCoachingProfile(user.coachingProfile);
  user.coachingProfile = normalized;
  return normalized;
}

function buildDashboardFromProfile(user) {
  const profile = ensureUserCoachingProfile(user);
  const logs = profile.interactionLog || [];
  const latest = logs[logs.length - 1];
  const previous = logs[logs.length - 2];
  const trends = {};
  for (const key of CATEGORY_KEYS) {
    const latestVal = latest ? latest.categoryScores[key] : profile.categoryScores[key];
    const previousVal = previous ? previous.categoryScores[key] : latestVal;
    trends[key] = latestVal - previousVal;
  }

  return {
    generatedAt: newIsoNow(),
    retentionDays: PROFILE_RETENTION_DAYS,
    profile: {
      traits: profile.traits,
      strengths: profile.strengths,
      developmentAreas: profile.developmentAreas,
      personality: profile.personality,
      categoryScores: profile.categoryScores,
      categoryTrends: trends,
      interventions: profile.interventions,
      interactionLog: logs.slice(-25).reverse(),
      metrics: profile.metrics,
      feedbackLoop: profile.feedbackLoop,
      updatedAt: profile.updatedAt,
    },
  };
}

function publicPracticePlan(plan) {
  const normalized = normalizePracticePlanRecord(plan);
  const persona = getStakeholderPersona(normalized.stakeholderPersonaId);
  const playbook = getDomainPlaybook(normalized.domainId);
  return {
    ...normalized,
    stakeholderPersona: persona
      ? {
          id: persona.id,
          name: persona.name,
          archetype: persona.archetype,
        }
      : null,
    playbook: playbook
      ? {
          id: playbook.id,
          name: playbook.name,
        }
      : null,
  };
}

function publicUser(user) {
  const profile = ensureUserCoachingProfile(user);
  const persona = getStakeholderPersona(user.selectedStakeholderPersonaId);
  const playbook = getDomainPlaybook(user.selectedPlaybookId);
  const practicePlans = normalizePracticePlans(user.practicePlans);
  user.practicePlans = practicePlans;
  syncUserNudges(user);
  return {
    id: user.id,
    username: user.username,
    email: user.email,
    displayName: user.displayName,
    role: user.role,
    focusAreas: user.focusAreas,
    coachingTone: normalizeCoachingTone(user.coachingTone, DEFAULT_COACHING_TONE),
    coachStyleWeights: normalizeCoachStyleWeights(
      user.coachStyleWeights,
      user.coachStyle,
    ),
    supervisorMode: Boolean(user.supervisorMode),
    selectedStakeholderPersonaId: normalizeStakeholderPersonaId(
      user.selectedStakeholderPersonaId,
      "",
    ),
    selectedPlaybookId: normalizePlaybookId(user.selectedPlaybookId, ""),
    selectedStakeholderPersona: persona
      ? {
          id: persona.id,
          name: persona.name,
          archetype: persona.archetype,
          pressureStyle: persona.pressureStyle,
        }
      : null,
    selectedPlaybook: playbook
      ? {
          id: playbook.id,
          name: playbook.name,
          context: playbook.context,
        }
      : null,
    practicePlansSummary: getPracticePlanSummary(practicePlans),
    profileUpdatedAt: profile.updatedAt,
    mustChangePassword: Boolean(user.mustChangePassword),
    isActive: Boolean(user.isActive !== false),
    nudgeSummary: getNudgeSummary(user),
    isAdmin: isAdminUser(user),
    createdAt: user.createdAt,
    updatedAt: user.updatedAt,
  };
}

function setCookie(res, cookieValue) {
  const current = res.getHeader("Set-Cookie");
  if (!current) {
    res.setHeader("Set-Cookie", cookieValue);
    return;
  }
  if (Array.isArray(current)) {
    res.setHeader("Set-Cookie", [...current, cookieValue]);
    return;
  }
  res.setHeader("Set-Cookie", [current, cookieValue]);
}

function parseCookieHeader(cookieHeader) {
  const cookies = {};
  const raw = String(cookieHeader || "");
  const parts = raw.split(";");
  for (const item of parts) {
    const [k, ...rest] = item.split("=");
    if (!k) continue;
    cookies[k.trim()] = decodeURIComponent(rest.join("=").trim());
  }
  return cookies;
}

function createSession(res, userId) {
  const sessionId = crypto.randomBytes(24).toString("base64url");
  sessions.set(sessionId, {
    userId,
    expiresAt: Date.now() + SESSION_TTL_MS,
  });
  setCookie(
    res,
    `${SESSION_COOKIE}=${sessionId}; HttpOnly; Path=/; Max-Age=${Math.floor(
      SESSION_TTL_MS / 1000,
    )}; SameSite=Lax`,
  );
}

function clearSession(req, res) {
  const cookies = parseCookieHeader(req.headers.cookie);
  const sessionId = cookies[SESSION_COOKIE];
  if (sessionId) sessions.delete(sessionId);
  setCookie(
    res,
    `${SESSION_COOKIE}=; HttpOnly; Path=/; Max-Age=0; SameSite=Lax`,
  );
}

function getAuthenticatedUser(req) {
  const cookies = parseCookieHeader(req.headers.cookie);
  const sessionId = cookies[SESSION_COOKIE];
  if (!sessionId) return null;
  const session = sessions.get(sessionId);
  if (!session) return null;
  if (session.expiresAt <= Date.now()) {
    sessions.delete(sessionId);
    return null;
  }
  const user = userStore.users.find((u) => u.id === session.userId);
  if (!user) {
    sessions.delete(sessionId);
    return null;
  }
  if (user.isActive === false) {
    sessions.delete(sessionId);
    return null;
  }
  syncUserNudges(user);
  session.expiresAt = Date.now() + SESSION_TTL_MS;
  sessions.set(sessionId, session);
  return user;
}

function isAdminUser(user) {
  if (!user || typeof user !== "object") return false;
  const username = normalizeUsername(user.username || "");
  const role = sanitizeText(user.role, 80).toLowerCase();
  return username === DEFAULT_ADMIN_USERNAME || role.includes("admin");
}

function buildUserContext(user) {
  const profile = ensureUserCoachingProfile(user);
  const appSettings = getAppSettings();
  const coachingTone = normalizeCoachingTone(
    user.coachingTone,
    DEFAULT_COACHING_TONE,
  );
  const weights = normalizeCoachStyleWeights(
    user.coachStyleWeights,
    user.coachStyle,
  );
  const supervisorMode = Boolean(user.supervisorMode && appSettings.supervisorModeEnabled);
  const persona = getStakeholderPersona(user.selectedStakeholderPersonaId);
  const playbook = getDomainPlaybook(user.selectedPlaybookId);
  const practicePlans = normalizePracticePlans(user.practicePlans);
  const practiceSummary = getPracticePlanSummary(practicePlans);
  const activePracticePlans = practicePlans.filter((plan) => plan.status === "active");
  const dominantStyles = getDominantStyles(weights).map(
    (key) => COACH_STYLE_CONFIG[key].label,
  );
  const lines = [
    "Leader profile:",
    `- Name: ${user.displayName || "Unknown"}`,
    `- Role: ${user.role || "Not specified"}`,
    `- Focus areas: ${user.focusAreas || "Not specified"}`,
    `- Strengths: ${profile.strengths.join(", ") || "None captured yet"}`,
    `- Development areas: ${profile.developmentAreas.join(", ") || "None captured yet"}`,
    `- Current category scores (1-10): decisionSpeed=${profile.categoryScores.decisionSpeed}, delegation=${profile.categoryScores.delegation}, conversationQuality=${profile.categoryScores.conversationQuality}, adaptability=${profile.categoryScores.adaptability}`,
    `- Feedback loop status: totalInteractions=${profile.feedbackLoop.totalInteractions}, analyzedInteractions=${profile.feedbackLoop.analyzedInteractions}, fallbackInteractions=${profile.feedbackLoop.fallbackInteractions}, lastAssessmentSource=${profile.feedbackLoop.lastAssessmentSource || "none"}`,
    `- Big Five (1-10): openness=${profile.personality.bigFive.openness}, conscientiousness=${profile.personality.bigFive.conscientiousness}, extraversion=${profile.personality.bigFive.extraversion}, agreeableness=${profile.personality.bigFive.agreeableness}, neuroticism=${profile.personality.bigFive.neuroticism}`,
    `- MBTI: ${profile.personality.mbti}`,
    `- DISC: ${profile.personality.disc}`,
    `- Supervisor mode: ${supervisorMode ? "enabled" : "disabled"}`,
    persona
      ? `- Selected stakeholder persona: ${persona.name} (${persona.archetype}). Pressure style: ${persona.pressureStyle}. Priorities: ${persona.priorities.join(", ")}. Guidance: ${persona.guidance}`
      : "- Selected stakeholder persona: none",
    playbook
      ? `- Selected domain playbook: ${playbook.name}. Context: ${playbook.context}. Core moves: ${playbook.coreMoves.join("; ")}. Pitfalls: ${playbook.pitfalls.join("; ")}`
      : "- Selected domain playbook: none",
    `- Practice plan summary: total=${practiceSummary.total}, active=${practiceSummary.active}, due=${practiceSummary.due}, completed=${practiceSummary.completed}`,
    ...activePracticePlans.slice(0, 3).map((plan, index) => {
      const playbookName =
        getDomainPlaybook(plan.domainId)?.name || "No playbook selected";
      const personaName =
        getStakeholderPersona(plan.stakeholderPersonaId)?.name || "No persona selected";
      const actions =
        Array.isArray(plan.actions) && plan.actions.length
          ? plan.actions.join("; ")
          : "No actions specified";
      return `- Active practice plan ${index + 1}: ${plan.title}. Objective: ${plan.objective || "Not specified"}. Cadence: ${plan.cadence}. Next due: ${plan.nextDueAt || "unscheduled"}. Playbook: ${playbookName}. Persona: ${personaName}. Actions: ${actions}`;
    }),
    `- Conversational coaching tone preference: ${coachingTone}`,
    "- Coaching style intensity profile (1=low, 5=high):",
    ...STYLE_KEYS.map(
      (key) =>
        `  - ${COACH_STYLE_CONFIG[key].label}: ${weights[key]}/5 (${COACH_STYLE_CONFIG[key].summary})`,
    ),
    `- Dominant coaching style(s): ${dominantStyles.join(", ")}`,
    "- Apply style instructions proportionate to the profile above:",
    ...STYLE_KEYS.map((key) => {
      const weight = weights[key];
      let emphasis = "low";
      if (weight >= 5) emphasis = "very high";
      else if (weight === 4) emphasis = "high";
      else if (weight === 3) emphasis = "balanced";
      else if (weight === 2) emphasis = "light";
      return `  - ${COACH_STYLE_CONFIG[key].label} (${emphasis} emphasis): ${COACH_STYLE_CONFIG[key].instruction}`;
    }),
    `- Apply tone preference: ${COACH_TONE_CONFIG[coachingTone]}`,
    supervisorMode
      ? "- Supervisor mode behavior: be more exacting and accountability-led. Challenge weak assumptions, insist on evidence, and always land one explicit commitment with owner and deadline."
      : "- Supervisor mode behavior: standard coaching stance.",
    "- Keep flow conversational. Ask one question, then wait. Avoid list-heavy outputs.",
  ];
  return lines.join("\n");
}

function getValidatedApiKey() {
  const apiKey = (process.env.OPENAI_API_KEY || "").trim();
  if (!apiKey) {
    return {
      ok: false,
      error: "OPENAI_API_KEY is not set. Export it before starting the server.",
    };
  }
  if (
    apiKey === "your_api_key" ||
    apiKey.includes("...") ||
    apiKey.includes("<") ||
    apiKey.includes(">") ||
    apiKey.toLowerCase().includes("replace_with")
  ) {
    return {
      ok: false,
      error:
        "OPENAI_API_KEY looks like a placeholder. Use a real Platform API key (starts with sk- or sk-proj-) with no template text.",
    };
  }
  return { ok: true, apiKey };
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function createOpenAIChatCompletion(apiKey, requestBody) {
  let lastError = null;
  const attempts = 3;

  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const res = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(requestBody),
        signal: AbortSignal.timeout(30000),
      });

      const payload = await res.json().catch(() => ({}));
      if (!res.ok) {
        const err = new Error(
          payload?.error?.message || "OpenAI API request failed.",
        );
        err.status = res.status;
        throw err;
      }

      return payload;
    } catch (err) {
      lastError = err;
      const status = Number(err?.status || 0);
      const retryable =
        !status || status >= 500 || status === 429 || err?.name === "TimeoutError";
      if (attempt < attempts && retryable) {
        await sleep(300 * attempt);
        continue;
      }
      break;
    }
  }

  throw lastError || new Error("OpenAI request failed.");
}

async function createOpenAISpeech(apiKey, requestBody) {
  let lastError = null;
  const attempts = 3;

  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const res = await fetch("https://api.openai.com/v1/audio/speech", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(requestBody),
        signal: AbortSignal.timeout(45000),
      });

      if (!res.ok) {
        const payload = await res.json().catch(() => ({}));
        const err = new Error(
          payload?.error?.message || "OpenAI text-to-speech request failed.",
        );
        err.status = res.status;
        throw err;
      }

      const raw = await res.arrayBuffer();
      return Buffer.from(raw);
    } catch (err) {
      lastError = err;
      const status = Number(err?.status || 0);
      const retryable =
        !status || status >= 500 || status === 429 || err?.name === "TimeoutError";
      if (attempt < attempts && retryable) {
        await sleep(300 * attempt);
        continue;
      }
      break;
    }
  }

  throw lastError || new Error("OpenAI text-to-speech request failed.");
}

async function persistUsers() {
  const payload = JSON.stringify(userStore, null, 2);
  persistQueue = persistQueue.then(() => writeFile(USERS_FILE, payload, "utf8"));
  return persistQueue;
}

function findUserByIdentifier(value) {
  const identifier = String(value || "")
    .trim()
    .toLowerCase();
  if (!identifier) return null;
  return (
    userStore.users.find(
      (u) => u.username === identifier || (u.email && u.email === identifier),
    ) || null
  );
}

function makeUniqueUsername(base, taken) {
  const cleaned = normalizeUsername(base)
    .replace(/[^a-z0-9_-]/g, "")
    .slice(0, 32);
  const seed = cleaned || "user";
  let candidate = seed.slice(0, 32);
  let index = 1;
  while (!candidate || taken.has(candidate)) {
    const suffix = String(index++);
    candidate = `${seed.slice(0, Math.max(1, 32 - suffix.length - 1))}_${suffix}`;
  }
  taken.add(candidate);
  return candidate;
}

function normalizeUsers() {
  const taken = new Set();
  let changed = false;

  for (const user of userStore.users) {
    const originalUsername = user.username;
    const normalizedExisting = normalizeUsername(originalUsername);
    const nextUsername =
      isValidUsername(normalizedExisting) && !taken.has(normalizedExisting)
        ? normalizedExisting
        : makeUniqueUsername(
            (user.email || "").split("@")[0] ||
              `user_${String(user.id || "").slice(0, 8)}`,
            taken,
          );

    if (!taken.has(nextUsername)) {
      taken.add(nextUsername);
    }
    if (originalUsername !== nextUsername) {
      user.username = nextUsername;
      changed = true;
    }

    const nextEmail = normalizeEmail(user.email || "");
    if (nextEmail !== user.email) {
      user.email = nextEmail;
      changed = true;
    }

    if (typeof user.mustChangePassword !== "boolean") {
      user.mustChangePassword = false;
      changed = true;
    }

    if (typeof user.isActive !== "boolean") {
      user.isActive = true;
      changed = true;
    }

    if (!user.displayName) {
      user.displayName = user.username;
      changed = true;
    }

    if (!user.role) {
      user.role = "";
      changed = true;
    }

    if (!user.focusAreas) {
      user.focusAreas = "";
      changed = true;
    }

    const nextTone = normalizeCoachingTone(
      user.coachingTone,
      DEFAULT_COACHING_TONE,
    );
    if (user.coachingTone !== nextTone) {
      user.coachingTone = nextTone;
      changed = true;
    }

    const normalizedWeights = normalizeCoachStyleWeights(
      user.coachStyleWeights,
      user.coachStyle,
    );
    if (!coachStyleWeightsEqual(user.coachStyleWeights || {}, normalizedWeights)) {
      user.coachStyleWeights = normalizedWeights;
      changed = true;
    }
    if (Object.hasOwn(user, "coachStyle")) {
      delete user.coachStyle;
      changed = true;
    }

    const supervisorMode = Boolean(user.supervisorMode);
    if (user.supervisorMode !== supervisorMode) {
      user.supervisorMode = supervisorMode;
      changed = true;
    }

    const normalizedPersonaId = normalizeStakeholderPersonaId(
      user.selectedStakeholderPersonaId,
      "",
    );
    if (user.selectedStakeholderPersonaId !== normalizedPersonaId) {
      user.selectedStakeholderPersonaId = normalizedPersonaId;
      changed = true;
    }

    const normalizedPlaybookId = normalizePlaybookId(user.selectedPlaybookId, "");
    if (user.selectedPlaybookId !== normalizedPlaybookId) {
      user.selectedPlaybookId = normalizedPlaybookId;
      changed = true;
    }

    const normalizedPracticePlans = normalizePracticePlans(user.practicePlans);
    if (
      JSON.stringify(user.practicePlans || []) !== JSON.stringify(normalizedPracticePlans)
    ) {
      user.practicePlans = normalizedPracticePlans;
      changed = true;
    }

    const normalizedProfile = normalizeCoachingProfile(user.coachingProfile);
    if (
      JSON.stringify(user.coachingProfile || {}) !==
      JSON.stringify(normalizedProfile)
    ) {
      user.coachingProfile = normalizedProfile;
      changed = true;
    }

    const normalizedNudges = normalizeNudges(user.nudges);
    if (JSON.stringify(user.nudges || []) !== JSON.stringify(normalizedNudges)) {
      user.nudges = normalizedNudges;
      changed = true;
    }
    if (syncUserNudges(user)) {
      changed = true;
    }

    if (!user.createdAt) {
      user.createdAt = new Date().toISOString();
      changed = true;
    }

    if (!user.updatedAt) {
      user.updatedAt = user.createdAt;
      changed = true;
    }
  }

  const normalizedSettings = normalizeAppSettings(userStore.appSettings);
  if (
    JSON.stringify(userStore.appSettings || {}) !==
    JSON.stringify(normalizedSettings)
  ) {
    userStore.appSettings = normalizedSettings;
    changed = true;
  }

  return changed;
}

async function ensureDefaultAdminUser() {
  if (userStore.users.some((u) => u.username === DEFAULT_ADMIN_USERNAME)) {
    return false;
  }

  const now = new Date().toISOString();
  userStore.users.push({
    id: crypto.randomUUID(),
    username: DEFAULT_ADMIN_USERNAME,
    email: "",
    passwordHash: hashPassword(DEFAULT_ADMIN_PASSWORD),
    displayName: "Administrator",
    role: "Administrator",
    focusAreas: "leadership decisions,difficult conversations",
    coachingTone: DEFAULT_COACHING_TONE,
    coachStyleWeights: { ...DEFAULT_COACH_STYLE_WEIGHTS },
    supervisorMode: false,
    selectedStakeholderPersonaId: "",
    selectedPlaybookId: "",
    practicePlans: [],
    coachingProfile: createDefaultCoachingProfile(),
    nudges: [],
    mustChangePassword: false,
    isActive: true,
    createdAt: now,
    updatedAt: now,
  });
  await persistUsers();
  return true;
}

async function loadUsers() {
  await mkdir(DATA_DIR, { recursive: true });
  try {
    const raw = await readFile(USERS_FILE, "utf8");
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.users)) {
      userStore = { users: [], appSettings: { ...DEFAULT_APP_SETTINGS } };
      await persistUsers();
      return;
    }
    userStore = parsed;
  } catch {
    userStore = { users: [], appSettings: { ...DEFAULT_APP_SETTINGS } };
    await persistUsers();
  }

  const changed = normalizeUsers();
  if (changed) {
    await persistUsers();
  }

  await ensureDefaultAdminUser();
}

function unauthorized(res) {
  sendJson(res, 401, { error: "Please log in first." });
}

async function handleRegister(req, res) {
  try {
    const appSettings = getAppSettings();
    if (appSettings.registrationDisabled) {
      sendJson(res, 403, {
        error: "New account creation is disabled. Contact an administrator.",
      });
      return;
    }

    const body = await parseBody(req);
    const username = normalizeUsername(body.username);
    const email = normalizeEmail(body.email);
    const password = String(body.password || "");
    const displayName = sanitizeText(body.displayName, 80);
    const role = sanitizeText(body.role, 120);
    const focusAreas = sanitizeText(body.focusAreas, 200);
    const coachingTone = normalizeCoachingTone(
      body.coachingTone,
      DEFAULT_COACHING_TONE,
    );
    const coachStyleWeights = normalizeCoachStyleWeights(
      body.coachStyleWeights,
      body.coachStyle,
    );

    if (!isValidUsername(username)) {
      sendJson(res, 400, {
        error:
          "Username must be 3-32 characters using lowercase letters, numbers, underscores, or hyphens.",
      });
      return;
    }
    if (email && !email.includes("@")) {
      sendJson(res, 400, { error: "Enter a valid email address." });
      return;
    }
    if (password.length < 10) {
      sendJson(res, 400, { error: "Password must be at least 10 characters." });
      return;
    }
    if (findUserByIdentifier(username)) {
      sendJson(res, 409, { error: "That username is already taken." });
      return;
    }
    if (email && userStore.users.find((u) => u.email && u.email === email)) {
      sendJson(res, 409, { error: "An account with that email already exists." });
      return;
    }

    const now = new Date().toISOString();
    const user = {
      id: crypto.randomUUID(),
      username,
      email,
      passwordHash: hashPassword(password),
      displayName: displayName || username,
      role,
      focusAreas,
      coachingTone,
      coachStyleWeights,
      supervisorMode: false,
      selectedStakeholderPersonaId: "",
      selectedPlaybookId: "",
      practicePlans: [],
      coachingProfile: createDefaultCoachingProfile(),
      nudges: [],
      mustChangePassword: false,
      isActive: true,
      createdAt: now,
      updatedAt: now,
    };

    userStore.users.push(user);
    await persistUsers();
    createSession(res, user.id);
    sendJson(res, 201, {
      user: publicUser(user),
      settings: getAppSettings(),
    });
  } catch (err) {
    sendJson(res, 400, { error: err.message || "Invalid register request." });
  }
}

async function handleLogin(req, res) {
  try {
    const body = await parseBody(req);
    const identifier = normalizeUsername(
      body.identifier ?? body.username ?? body.email,
    );
    const password = String(body.password || "");

    const user = findUserByIdentifier(identifier);
    if (!user || !verifyPassword(password, user.passwordHash)) {
      sendJson(res, 401, { error: "Invalid username/email or password." });
      return;
    }
    if (user.isActive === false) {
      sendJson(res, 403, {
        error: "Account is deactivated. Contact an administrator.",
      });
      return;
    }

    createSession(res, user.id);
    sendJson(res, 200, {
      user: publicUser(user),
      settings: getAppSettings(),
    });
  } catch (err) {
    sendJson(res, 400, { error: err.message || "Invalid login request." });
  }
}

async function handlePasswordChange(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }

  try {
    const body = await parseBody(req);
    const currentPassword = String(body.currentPassword || "");
    const newPassword = String(body.newPassword || "");

    if (!verifyPassword(currentPassword, user.passwordHash)) {
      sendJson(res, 401, { error: "Current password is incorrect." });
      return;
    }
    if (newPassword.length < 10) {
      sendJson(res, 400, {
        error: "New password must be at least 10 characters.",
      });
      return;
    }
    if (verifyPassword(newPassword, user.passwordHash)) {
      sendJson(res, 400, {
        error: "New password must be different from current password.",
      });
      return;
    }

    user.passwordHash = hashPassword(newPassword);
    user.mustChangePassword = false;
    user.updatedAt = new Date().toISOString();
    await persistUsers();
    sendJson(res, 200, { user: publicUser(user) });
  } catch (err) {
    sendJson(res, 400, { error: err.message || "Invalid password request." });
  }
}

function handleLogout(req, res) {
  clearSession(req, res);
  sendJson(res, 200, { ok: true });
}

function handleMe(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    sendJson(res, 200, {
      authenticated: false,
      settings: getAppSettings(),
    });
    return;
  }
  sendJson(res, 200, {
    authenticated: true,
    user: publicUser(user),
    settings: getAppSettings(),
  });
}

function handleDashboard(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }
  sendJson(res, 200, {
    user: publicUser(user),
    dashboard: buildDashboardFromProfile(user),
  });
}

function incrementCount(map, key) {
  const safeKey = sanitizeText(key, 120) || "Unknown";
  map.set(safeKey, (map.get(safeKey) || 0) + 1);
}

function topCounts(map, limit = 5) {
  return Array.from(map.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, limit)
    .map(([label, count]) => ({ label, count }));
}

function buildAdminDashboard() {
  const generatedAt = newIsoNow();
  const now = Date.now();
  const activeCutoff = now - 30 * 24 * 60 * 60 * 1000;
  let totalInteractions = 0;
  let totalSelfReports = 0;
  let totalMeetingNotes = 0;
  let totalFeedbackSubmissions = 0;
  let feedbackScoreAccum = 0;
  let feedbackScoreCount = 0;
  let totalDueNudges = 0;
  let totalScheduledNudges = 0;
  let activeUsers30d = 0;
  let totalPracticePlans = 0;
  let totalActivePracticePlans = 0;
  let usersInSupervisorMode = 0;
  let usersWithPersonaSelection = 0;
  let usersWithPlaybookSelection = 0;
  let avgScoreAccum = 0;
  const globalUseCases = new Map();

  const users = userStore.users.map((user) => {
    syncUserNudges(user);
    const profile = ensureUserCoachingProfile(user);
    const logs = Array.isArray(profile.interactionLog) ? profile.interactionLog : [];
    const nudgeSummary = getNudgeSummary(user);
    const practicePlans = normalizePracticePlans(user.practicePlans);
    const practicePlansSummary = getPracticePlanSummary(practicePlans);
    user.practicePlans = practicePlans;
    const selectedStakeholderPersona = getStakeholderPersona(
      user.selectedStakeholderPersonaId,
    );
    const selectedPlaybook = getDomainPlaybook(user.selectedPlaybookId);
    if (user.supervisorMode) usersInSupervisorMode += 1;
    if (selectedStakeholderPersona) usersWithPersonaSelection += 1;
    if (selectedPlaybook) usersWithPlaybookSelection += 1;
    const lastInteractionAt =
      logs[logs.length - 1]?.timestamp || profile.updatedAt || user.updatedAt || "";
    const lastMs = toMs(lastInteractionAt);
    if (lastMs >= activeCutoff) {
      activeUsers30d += 1;
    }

    const useCases = new Map();
    let selfReportCount = 0;
    let meetingNotesCount = 0;
    let fallbackAssessments = 0;
    let busySessions = 0;
    let standardSessions = 0;
    let feedbackSubmissionCount = 0;
    let feedbackUserScoreAccum = 0;
    let feedbackUserScoreCount = 0;
    const recentSelfReports = [];
    const recentCoacheeFeedback = [];

    for (const entry of logs) {
      const useCase =
        sanitizeText(entry?.useCase, 120) ||
        inferUseCaseFromText(entry?.summary, "General coaching");
      incrementCount(useCases, useCase);
      incrementCount(globalUseCases, useCase);

      const isBusy = normalizeSessionPace(entry?.sessionPace) === "busy";
      if (isBusy) busySessions += 1;
      else standardSessions += 1;

      if (entry?.assessmentMode === "fallback") {
        fallbackAssessments += 1;
      }
      if (entry?.evidenceFlags?.selfReportProvided) {
        selfReportCount += 1;
      }
      if (entry?.evidenceFlags?.meetingNotesProvided) {
        meetingNotesCount += 1;
      }
      const feedbackSnippet = sanitizeText(entry?.selfReportSnippet, 220);
      if (feedbackSnippet) {
        recentSelfReports.push({
          timestamp: entry.timestamp || "",
          text: feedbackSnippet,
        });
      }
      const coacheeFeedback = normalizeInteractionFeedback(entry?.coacheeFeedback);
      if (coacheeFeedback) {
        feedbackSubmissionCount += 1;
        const scoreParts = [
          coacheeFeedback.helpfulness,
          coacheeFeedback.clarity,
          coacheeFeedback.confidence,
        ].filter((value) => Number.isFinite(value));
        if (scoreParts.length) {
          const average = Number(
            (
              scoreParts.reduce((sum, value) => sum + value, 0) /
              scoreParts.length
            ).toFixed(1),
          );
          feedbackUserScoreAccum += average;
          feedbackUserScoreCount += 1;
          feedbackScoreAccum += average;
          feedbackScoreCount += 1;
        }
        recentCoacheeFeedback.push({
          timestamp: coacheeFeedback.submittedAt || entry.timestamp || "",
          helpfulness: coacheeFeedback.helpfulness,
          clarity: coacheeFeedback.clarity,
          confidence: coacheeFeedback.confidence,
          comment: coacheeFeedback.comment,
        });
      }
    }

    totalInteractions += logs.length;
    totalSelfReports += selfReportCount;
    totalMeetingNotes += meetingNotesCount;
    totalFeedbackSubmissions += feedbackSubmissionCount;
    totalDueNudges += nudgeSummary.due;
    totalScheduledNudges += nudgeSummary.scheduled;
    totalPracticePlans += practicePlansSummary.total;
    totalActivePracticePlans += practicePlansSummary.active;
    avgScoreAccum += Number(profile.metrics?.averageCategoryScore || 0);
    const averageFeedbackScore =
      feedbackUserScoreCount > 0
        ? Number((feedbackUserScoreAccum / feedbackUserScoreCount).toFixed(2))
        : null;

    return {
      id: user.id,
      username: user.username,
      displayName: user.displayName,
      role: user.role,
      email: user.email,
      focusAreas: user.focusAreas,
      isActive: Boolean(user.isActive !== false),
      mustChangePassword: Boolean(user.mustChangePassword),
      createdAt: user.createdAt,
      updatedAt: user.updatedAt,
      profileUpdatedAt: profile.updatedAt,
      nudgeSummary,
      supervisorMode: Boolean(user.supervisorMode),
      selectedStakeholderPersona: selectedStakeholderPersona
        ? {
            id: selectedStakeholderPersona.id,
            name: selectedStakeholderPersona.name,
          }
        : null,
      selectedPlaybook: selectedPlaybook
        ? {
            id: selectedPlaybook.id,
            name: selectedPlaybook.name,
          }
        : null,
      practicePlansSummary,
      usage: {
        totalInteractions: logs.length,
        lastInteractionAt,
        sessionPace: {
          busy: busySessions,
          standard: standardSessions,
        },
        selfReportCount,
        meetingNotesCount,
        fallbackAssessments,
        feedbackSubmissionCount,
        averageFeedbackScore,
        topUseCases: topCounts(useCases, 3),
      },
      usedFor: topCounts(useCases, 5),
      personality: profile.personality,
      categoryScores: profile.categoryScores,
      metrics: profile.metrics,
      feedback: {
        recentSelfReports: recentSelfReports.slice(-3).reverse(),
        recentCoacheeFeedback: recentCoacheeFeedback.slice(-3).reverse(),
        recentInteractionSummaries: logs
          .slice(-3)
          .reverse()
          .map((entry) => ({
            timestamp: entry.timestamp,
            summary: sanitizeText(entry.summary, 220),
            useCase:
              sanitizeText(entry.useCase, 120) ||
              inferUseCaseFromText(entry.summary, "General coaching"),
          })),
      },
    };
  });

  const totalUsers = users.length;
  const avgCategoryScore =
    totalUsers > 0
      ? Number((avgScoreAccum / Math.max(1, totalUsers)).toFixed(2))
      : 0;
  const avgFeedbackScore =
    feedbackScoreCount > 0
      ? Number((feedbackScoreAccum / feedbackScoreCount).toFixed(2))
      : null;
  const activeAccounts = users.filter((item) => item.isActive).length;
  const inactiveAccounts = totalUsers - activeAccounts;

  return {
    generatedAt,
    summary: {
      totalUsers,
      activeAccounts,
      inactiveAccounts,
      activeUsers30d,
      totalInteractions,
      totalSelfReports,
      totalMeetingNotes,
      totalFeedbackSubmissions,
      averageFeedbackScore: avgFeedbackScore,
      totalNudgesDue: totalDueNudges,
      totalNudgesScheduled: totalScheduledNudges,
      totalPracticePlans,
      totalActivePracticePlans,
      usersInSupervisorMode,
      usersWithPersonaSelection,
      usersWithPlaybookSelection,
      averageCategoryScore: avgCategoryScore,
      topUseCases: topCounts(globalUseCases, 6),
      registrationDisabled: getAppSettings().registrationDisabled,
      supervisorModeEnabled: getAppSettings().supervisorModeEnabled,
    },
    users: users.sort((a, b) => toMs(b.usage.lastInteractionAt) - toMs(a.usage.lastInteractionAt)),
  };
}

function handleAdminDashboard(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }
  if (!isAdminUser(user)) {
    sendJson(res, 403, { error: "Admin access required." });
    return;
  }
  sendJson(res, 200, {
    user: publicUser(user),
    settings: getAppSettings(),
    dashboard: buildAdminDashboard(),
  });
}

async function handleAdminSettingsUpdate(req, res) {
  const user = requireAdmin(req, res);
  if (!user) return;
  try {
    const body = await parseBody(req);
    const current = getAppSettings();
    const next = normalizeAppSettings({
      ...current,
      registrationDisabled: body.registrationDisabled,
      supervisorModeEnabled: body.supervisorModeEnabled,
    });
    userStore.appSettings = next;
    if (!next.supervisorModeEnabled) {
      for (const item of userStore.users) {
        if (item.supervisorMode) {
          item.supervisorMode = false;
          item.updatedAt = newIsoNow();
        }
      }
    }
    await persistUsers();
    sendJson(res, 200, {
      ok: true,
      user: publicUser(user),
      settings: getAppSettings(),
      dashboard: buildAdminDashboard(),
    });
  } catch (err) {
    sendJson(res, 400, { error: err.message || "Invalid admin settings request." });
  }
}

function clearSessionsForUser(userId) {
  for (const [sessionId, session] of sessions.entries()) {
    if (session?.userId === userId) {
      sessions.delete(sessionId);
    }
  }
}

function requireAdmin(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return null;
  }
  if (!isAdminUser(user)) {
    sendJson(res, 403, { error: "Admin access required." });
    return null;
  }
  return user;
}

async function handleAdminCreateUser(req, res) {
  const adminUser = requireAdmin(req, res);
  if (!adminUser) return;

  try {
    const body = await parseBody(req);
    const username = normalizeUsername(body.username);
    const email = normalizeEmail(body.email);
    const password = String(body.password || "");
    const displayName = sanitizeText(body.displayName, 80);
    const role = sanitizeText(body.role, 120);
    const focusAreas = sanitizeText(body.focusAreas, 200);
    const coachingTone = normalizeCoachingTone(
      body.coachingTone,
      DEFAULT_COACHING_TONE,
    );
    const coachStyleWeights = normalizeCoachStyleWeights(
      body.coachStyleWeights,
      body.coachStyle,
    );
    const mustChangePassword = body.mustChangePassword !== false;

    if (!isValidUsername(username)) {
      sendJson(res, 400, {
        error:
          "Username must be 3-32 characters using lowercase letters, numbers, underscores, or hyphens.",
      });
      return;
    }
    if (email && !email.includes("@")) {
      sendJson(res, 400, { error: "Enter a valid email address." });
      return;
    }
    if (password.length < 8) {
      sendJson(res, 400, {
        error: "Password must be at least 8 characters.",
      });
      return;
    }
    if (findUserByIdentifier(username)) {
      sendJson(res, 409, { error: "That username is already taken." });
      return;
    }
    if (email && userStore.users.find((u) => u.email && u.email === email)) {
      sendJson(res, 409, { error: "An account with that email already exists." });
      return;
    }

    const now = newIsoNow();
    const createdUser = {
      id: crypto.randomUUID(),
      username,
      email,
      passwordHash: hashPassword(password),
      displayName: displayName || username,
      role,
      focusAreas,
      coachingTone,
      coachStyleWeights,
      supervisorMode: false,
      selectedStakeholderPersonaId: "",
      selectedPlaybookId: "",
      practicePlans: [],
      coachingProfile: createDefaultCoachingProfile(),
      nudges: [],
      mustChangePassword,
      isActive: true,
      createdAt: now,
      updatedAt: now,
    };

    userStore.users.push(createdUser);
    await persistUsers();
    sendJson(res, 201, {
      ok: true,
      user: publicUser(adminUser),
      createdUser: publicUser(createdUser),
      settings: getAppSettings(),
      dashboard: buildAdminDashboard(),
    });
  } catch (err) {
    sendJson(res, 400, { error: err.message || "Invalid admin user request." });
  }
}

function resolveAdminAction(value) {
  const raw = sanitizeText(value, 60).toLowerCase();
  if (raw === "activate") return "activate";
  if (raw === "deactivate" || raw === "suspend") return "deactivate";
  if (
    raw === "force_password_reset" ||
    raw === "force-password-reset" ||
    raw === "forcepasswordreset"
  ) {
    return "force_password_reset";
  }
  if (
    raw === "reset_password" ||
    raw === "reset-password" ||
    raw === "resetpassword"
  ) {
    return "reset_password";
  }
  return "";
}

async function handleAdminUserAction(req, res, targetUserId) {
  const adminUser = requireAdmin(req, res);
  if (!adminUser) return;

  const targetId = sanitizeText(targetUserId, 80);
  const targetUser = userStore.users.find((item) => item.id === targetId);
  if (!targetUser) {
    sendJson(res, 404, { error: "User not found." });
    return;
  }

  try {
    const body = await parseBody(req);
    const action = resolveAdminAction(body.action);
    if (!action) {
      sendJson(res, 400, {
        error:
          "Invalid action. Use activate, deactivate, force_password_reset, or reset_password.",
      });
      return;
    }

    if (action === "deactivate") {
      if (targetUser.id === adminUser.id) {
        sendJson(res, 400, {
          error: "You cannot deactivate your own active admin session.",
        });
        return;
      }
      targetUser.isActive = false;
      clearSessionsForUser(targetUser.id);
    } else if (action === "activate") {
      targetUser.isActive = true;
    } else if (action === "force_password_reset") {
      targetUser.mustChangePassword = true;
      clearSessionsForUser(targetUser.id);
    } else if (action === "reset_password") {
      const newPassword = String(body.newPassword || "");
      if (newPassword.length < 10) {
        sendJson(res, 400, {
          error: "Temporary password must be at least 10 characters.",
        });
        return;
      }
      targetUser.passwordHash = hashPassword(newPassword);
      targetUser.mustChangePassword = true;
      targetUser.isActive = true;
      clearSessionsForUser(targetUser.id);
    }

    targetUser.updatedAt = newIsoNow();
    await persistUsers();
    sendJson(res, 200, {
      ok: true,
      action,
      user: publicUser(adminUser),
      targetUser: publicUser(targetUser),
      settings: getAppSettings(),
      dashboard: buildAdminDashboard(),
    });
  } catch (err) {
    sendJson(res, 400, { error: err.message || "Invalid admin action request." });
  }
}

function scoreFeedback(feedback) {
  if (!feedback || typeof feedback !== "object") return null;
  const parts = [feedback.helpfulness, feedback.clarity, feedback.confidence].filter(
    (value) => Number.isFinite(value),
  );
  if (!parts.length) return null;
  return Number((parts.reduce((sum, value) => sum + value, 0) / parts.length).toFixed(1));
}

function applyFeedbackToInterventions(profile, feedback) {
  const score = scoreFeedback(feedback);
  if (!Number.isFinite(score)) return;

  const now = newIsoNow();
  for (const intervention of profile.interventions) {
    if (!intervention || typeof intervention !== "object") continue;
    if (!["active", "improving"].includes(intervention.status)) continue;
    const adjusted = Math.round((Number(intervention.successScore || 5) * 0.65 + score * 0.35) * 10) / 10;
    intervention.successScore = clampScore10(adjusted, intervention.successScore || 5);
    if (intervention.successScore >= 9) {
      intervention.status = "succeeded";
    } else if (intervention.successScore >= 7 && intervention.status === "active") {
      intervention.status = "improving";
    }
    const feedbackComment = sanitizeText(feedback.comment, 140);
    if (feedbackComment) {
      const evidenceLine = sanitizeText(
        `Coachee feedback (${feedback.submittedAt || now}): ${feedbackComment}`,
        180,
      );
      if (evidenceLine) {
        intervention.evidence = sanitizeArray(
          [...(Array.isArray(intervention.evidence) ? intervention.evidence : []), evidenceLine],
          6,
          180,
        );
      }
    }
    intervention.updatedAt = now;
  }
}

async function handleFeedback(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }

  try {
    const body = await parseBody(req);
    const feedback = normalizeInteractionFeedback(body);
    if (!feedback) {
      sendJson(res, 400, {
        error: "Provide at least one feedback score or a comment.",
      });
      return;
    }
    const profile = ensureUserCoachingProfile(user);
    const logs = Array.isArray(profile.interactionLog) ? profile.interactionLog : [];
    const latest = logs[logs.length - 1];
    if (!latest) {
      sendJson(res, 400, {
        error: "No coaching interaction found yet. Send a coaching message first.",
      });
      return;
    }
    latest.coacheeFeedback = feedback;
    if (!latest.evidenceFlags || typeof latest.evidenceFlags !== "object") {
      latest.evidenceFlags = {
        meetingNotesProvided: false,
        selfReportProvided: false,
      };
    }
    if (feedback.comment && !latest.selfReportSnippet) {
      latest.selfReportSnippet = sanitizeText(feedback.comment, 240);
      latest.evidenceFlags.selfReportProvided = true;
    }
    applyFeedbackToInterventions(profile, feedback);
    profile.feedbackLoop = recomputeFeedbackLoop(profile);
    profile.metrics = computeProfileMetrics(profile);
    profile.updatedAt = newIsoNow();
    user.coachingProfile = normalizeCoachingProfile(profile);
    user.updatedAt = profile.updatedAt;
    await persistUsers();
    sendJson(res, 200, {
      ok: true,
      feedback: latest.coacheeFeedback,
      dashboard: buildDashboardFromProfile(user),
      user: publicUser(user),
    });
  } catch (err) {
    sendJson(res, 400, { error: err.message || "Invalid feedback request." });
  }
}

function parseNudgeDueAt(value) {
  const provided = sanitizeText(value, 60);
  if (!provided) return new Date(Date.now() + 48 * 60 * 60 * 1000).toISOString();
  const ms = toMs(provided);
  if (!ms) return "";
  return new Date(ms).toISOString();
}

async function handleNudgesList(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }
  const changed = syncUserNudges(user);
  if (changed) {
    user.updatedAt = newIsoNow();
    await persistUsers();
  }
  sendJson(res, 200, {
    nudges: normalizeNudges(user.nudges),
    summary: getNudgeSummary(user),
  });
}

async function handleNudgesCreate(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }

  try {
    syncUserNudges(user);
    const existing = normalizeNudges(user.nudges);
    if (existing.length >= MAX_NUDGES_PER_USER) {
      sendJson(res, 400, {
        error: `Maximum ${MAX_NUDGES_PER_USER} nudges reached. Complete or dismiss existing nudges first.`,
      });
      return;
    }

    const body = await parseBody(req);
    const dueAt = parseNudgeDueAt(body.dueAt);
    if (!dueAt) {
      sendJson(res, 400, { error: "Provide a valid nudge due date/time." });
      return;
    }
    const dueMs = toMs(dueAt);
    const nowIso = newIsoNow();
    const statusFromBody = normalizeNudgeStatus(body.status, "scheduled");
    const nudge = normalizeNudgeRecord({
      title: body.title,
      message: body.message,
      dueAt,
      status:
        statusFromBody === "scheduled" && dueMs <= Date.now()
          ? "due"
          : statusFromBody,
      createdAt: nowIso,
      updatedAt: nowIso,
    });

    user.nudges = normalizeNudges([...existing, nudge]);
    syncUserNudges(user);
    user.updatedAt = nowIso;
    await persistUsers();
    sendJson(res, 201, {
      ok: true,
      nudge,
      nudges: user.nudges,
      summary: getNudgeSummary(user),
      user: publicUser(user),
    });
  } catch (err) {
    sendJson(res, 400, { error: err.message || "Invalid nudge create request." });
  }
}

async function handleNudgeUpdate(req, res, nudgeId) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }

  try {
    syncUserNudges(user);
    const list = normalizeNudges(user.nudges);
    const targetId = sanitizeText(nudgeId, 80);
    const target = list.find((item) => item.id === targetId);
    if (!target) {
      sendJson(res, 404, { error: "Nudge not found." });
      return;
    }
    const body = await parseBody(req);
    if (Object.hasOwn(body, "title")) {
      target.title = sanitizeText(body.title, 120) || target.title;
    }
    if (Object.hasOwn(body, "message")) {
      target.message = sanitizeText(body.message, 320);
    }
    if (Object.hasOwn(body, "dueAt")) {
      const dueAt = parseNudgeDueAt(body.dueAt);
      if (!dueAt) {
        sendJson(res, 400, { error: "Provide a valid nudge due date/time." });
        return;
      }
      target.dueAt = dueAt;
    }
    if (Object.hasOwn(body, "status")) {
      target.status = normalizeNudgeStatus(body.status, target.status);
    }
    target.updatedAt = newIsoNow();

    const dueMs = toMs(target.dueAt);
    if (target.status === "scheduled" && dueMs && dueMs <= Date.now()) {
      target.status = "due";
    }
    if (
      target.status === "due" &&
      dueMs &&
      dueMs > Date.now() &&
      !Object.hasOwn(body, "status")
    ) {
      target.status = "scheduled";
    }

    user.nudges = normalizeNudges(list);
    syncUserNudges(user);
    user.updatedAt = newIsoNow();
    await persistUsers();
    sendJson(res, 200, {
      ok: true,
      nudge: target,
      nudges: user.nudges,
      summary: getNudgeSummary(user),
      user: publicUser(user),
    });
  } catch (err) {
    sendJson(res, 400, { error: err.message || "Invalid nudge update request." });
  }
}

function parsePracticePlanActions(value) {
  if (Array.isArray(value)) {
    return sanitizeArray(value, 8, 200);
  }
  const text = sanitizeText(value, 2000);
  if (!text) return [];
  const rawItems = text
    .split(/\n|;|,/g)
    .map((item) => item.trim())
    .filter(Boolean);
  return sanitizeArray(rawItems, 8, 200);
}

function handleStakeholderPersonas(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }
  sendJson(res, 200, {
    personas: STAKEHOLDER_PERSONAS,
  });
}

function handleDomainPlaybooks(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }
  sendJson(res, 200, {
    playbooks: DOMAIN_PLAYBOOKS,
  });
}

async function handlePracticePlansList(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }
  user.practicePlans = normalizePracticePlans(user.practicePlans);
  sendJson(res, 200, {
    plans: user.practicePlans.map((plan) => publicPracticePlan(plan)),
    summary: getPracticePlanSummary(user.practicePlans),
  });
}

async function handlePracticePlansCreate(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }

  try {
    const existing = normalizePracticePlans(user.practicePlans);
    if (existing.length >= MAX_PRACTICE_PLANS_PER_USER) {
      sendJson(res, 400, {
        error: `Maximum ${MAX_PRACTICE_PLANS_PER_USER} practice plans reached.`,
      });
      return;
    }

    const body = await parseBody(req);
    const title = sanitizeText(body.title, 140);
    if (!title) {
      sendJson(res, 400, { error: "Practice plan title is required." });
      return;
    }

    const now = newIsoNow();
    const plan = normalizePracticePlanRecord({
      title,
      objective: body.objective,
      domainId: body.domainId,
      stakeholderPersonaId: body.stakeholderPersonaId,
      actions: parsePracticePlanActions(body.actions),
      cadence: body.cadence,
      status: body.status || "active",
      nextDueAt: body.nextDueAt,
      notes: body.notes,
      createdAt: now,
      updatedAt: now,
    });

    user.practicePlans = normalizePracticePlans([plan, ...existing]);
    user.updatedAt = now;
    await persistUsers();
    sendJson(res, 201, {
      ok: true,
      plan: publicPracticePlan(plan),
      plans: user.practicePlans.map((item) => publicPracticePlan(item)),
      summary: getPracticePlanSummary(user.practicePlans),
      user: publicUser(user),
    });
  } catch (err) {
    sendJson(res, 400, { error: err.message || "Invalid practice plan request." });
  }
}

async function handlePracticePlanUpdate(req, res, planId) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }

  try {
    const plans = normalizePracticePlans(user.practicePlans);
    const targetId = sanitizeText(planId, 80);
    const index = plans.findIndex((item) => item.id === targetId);
    if (index < 0) {
      sendJson(res, 404, { error: "Practice plan not found." });
      return;
    }

    const body = await parseBody(req);
    const action = sanitizeText(body.action, 40).toLowerCase();
    const target = { ...plans[index] };
    const now = newIsoNow();

    if (action === "mark_practiced" || action === "complete_session") {
      target.lastPracticedAt = now;
      target.completedCount = Math.min(
        1000,
        Number(target.completedCount || 0) + 1,
      );
      if (target.status !== "completed") {
        target.status = "active";
      }
      if (body.nextDueAt) {
        target.nextDueAt = normalizeOptionalIsoDate(body.nextDueAt);
      }
    } else if (action === "complete") {
      target.status = "completed";
      target.lastPracticedAt = target.lastPracticedAt || now;
    } else if (action === "pause") {
      target.status = "paused";
    } else if (action === "activate") {
      target.status = "active";
    } else {
      if (Object.hasOwn(body, "title")) {
        target.title = sanitizeText(body.title, 140) || target.title;
      }
      if (Object.hasOwn(body, "objective")) {
        target.objective = sanitizeText(body.objective, 420);
      }
      if (Object.hasOwn(body, "domainId")) {
        target.domainId = normalizePlaybookId(body.domainId, target.domainId);
      }
      if (Object.hasOwn(body, "stakeholderPersonaId")) {
        target.stakeholderPersonaId = normalizeStakeholderPersonaId(
          body.stakeholderPersonaId,
          target.stakeholderPersonaId,
        );
      }
      if (Object.hasOwn(body, "actions")) {
        target.actions = parsePracticePlanActions(body.actions);
      }
      if (Object.hasOwn(body, "cadence")) {
        target.cadence = normalizePracticePlanCadence(body.cadence, target.cadence);
      }
      if (Object.hasOwn(body, "status")) {
        target.status = normalizePracticePlanStatus(body.status, target.status);
      }
      if (Object.hasOwn(body, "nextDueAt")) {
        target.nextDueAt = normalizeOptionalIsoDate(body.nextDueAt);
      }
      if (Object.hasOwn(body, "notes")) {
        target.notes = sanitizeText(body.notes, 420);
      }
    }

    target.updatedAt = now;
    plans[index] = normalizePracticePlanRecord(target);
    user.practicePlans = normalizePracticePlans(plans);
    user.updatedAt = now;
    await persistUsers();
    const savedPlan =
      user.practicePlans.find((item) => item.id === target.id) || target;
    sendJson(res, 200, {
      ok: true,
      plan: publicPracticePlan(savedPlan),
      plans: user.practicePlans.map((item) => publicPracticePlan(item)),
      summary: getPracticePlanSummary(user.practicePlans),
      user: publicUser(user),
    });
  } catch (err) {
    sendJson(res, 400, { error: err.message || "Invalid practice plan update." });
  }
}

async function handleProfileUpdate(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }

  try {
    const body = await parseBody(req);
    const appSettings = getAppSettings();
    const displayName = sanitizeText(body.displayName ?? user.displayName, 80);
    const role = sanitizeText(body.role ?? user.role, 120);
    const focusAreas = sanitizeText(body.focusAreas ?? user.focusAreas, 200);
    const coachingTone = normalizeCoachingTone(
      body.coachingTone ?? user.coachingTone,
      DEFAULT_COACHING_TONE,
    );
    const coachStyleWeights = normalizeCoachStyleWeights(
      body.coachStyleWeights ?? user.coachStyleWeights,
      body.coachStyle ?? user.coachStyle,
    );
    const selectedStakeholderPersonaId = normalizeStakeholderPersonaId(
      body.selectedStakeholderPersonaId ?? user.selectedStakeholderPersonaId,
      "",
    );
    const selectedPlaybookId = normalizePlaybookId(
      body.selectedPlaybookId ?? user.selectedPlaybookId,
      "",
    );
    const requestedSupervisorMode =
      typeof body.supervisorMode === "boolean"
        ? body.supervisorMode
        : Boolean(user.supervisorMode);
    const supervisorMode = appSettings.supervisorModeEnabled
      ? requestedSupervisorMode
      : false;

    user.displayName = displayName || user.displayName;
    user.role = role;
    user.focusAreas = focusAreas;
    user.coachingTone = coachingTone;
    user.coachStyleWeights = coachStyleWeights;
    user.selectedStakeholderPersonaId = selectedStakeholderPersonaId;
    user.selectedPlaybookId = selectedPlaybookId;
    user.supervisorMode = supervisorMode;
    user.updatedAt = new Date().toISOString();

    await persistUsers();
    sendJson(res, 200, { user: publicUser(user) });
  } catch (err) {
    sendJson(res, 400, { error: err.message || "Invalid profile update." });
  }
}

function extractJsonObject(text) {
  const raw = String(text || "").trim();
  if (!raw) return {};
  try {
    return JSON.parse(raw);
  } catch {
    const start = raw.indexOf("{");
    const end = raw.lastIndexOf("}");
    if (start >= 0 && end > start) {
      return JSON.parse(raw.slice(start, end + 1));
    }
    return {};
  }
}

function mergeInterventions(existing, incoming) {
  const byId = new Map();
  for (const item of existing) {
    byId.set(item.id, normalizeInterventionRecord(item));
  }
  for (const item of incoming) {
    const normalized = normalizeInterventionRecord(item);
    if (item?.id && byId.has(item.id)) {
      const current = byId.get(item.id);
      byId.set(item.id, {
        ...current,
        ...normalized,
        id: current.id,
        createdAt: current.createdAt,
        updatedAt: newIsoNow(),
      });
      continue;
    }
    byId.set(normalized.id, { ...normalized, updatedAt: newIsoNow() });
  }
  return Array.from(byId.values())
    .sort((a, b) => toMs(b.updatedAt) - toMs(a.updatedAt))
    .slice(0, 30);
}

async function analyzeAndUpdateCoacheeProfile({
  apiKey,
  user,
  userMessage,
  assistantReply,
  meetingNotes,
  selfReport,
  sessionPace,
}) {
  const profile = ensureUserCoachingProfile(user);
  const beforeScores = { ...profile.categoryScores };
  const now = newIsoNow();

  const analyzerInput = {
    timestamp: now,
    leaderProfile: {
      displayName: user.displayName || "",
      role: user.role || "",
      focusAreas: user.focusAreas || "",
    },
    currentProfile: profile,
    latestInteraction: {
      userMessage: sanitizeText(userMessage, 4000),
      assistantReply: sanitizeText(assistantReply, 5000),
      meetingNotes: sanitizeText(meetingNotes, 6000),
      selfReport: sanitizeText(selfReport, 3000),
    },
  };
  const payload = await createOpenAIChatCompletion(apiKey, {
    model: ANALYSIS_MODEL,
    temperature: 0.2,
    response_format: { type: "json_object" },
    messages: [
      { role: "system", content: PROFILE_ANALYZER_PROMPT },
      { role: "user", content: JSON.stringify(analyzerInput) },
    ],
  });
  const content = payload?.choices?.[0]?.message?.content || "{}";
  const analysis = extractJsonObject(content);

  profile.traits = sanitizeArray(analysis.traits, 12, 80);
  profile.strengths = sanitizeArray(analysis.strengths, 12, 120);
  profile.developmentAreas = sanitizeArray(analysis.developmentAreas, 12, 120);
  profile.personality = {
    bigFive: normalizeScoreMap(
      analysis?.personality?.bigFive,
      BIG_FIVE_KEYS,
      DEFAULT_BIG_FIVE,
    ),
    mbti: sanitizeText(analysis?.personality?.mbti || "Unknown", 10) || "Unknown",
    disc: sanitizeText(analysis?.personality?.disc || "Unknown", 10) || "Unknown",
  };
  profile.categoryScores = normalizeScoreMap(
    analysis.categoryScores,
    CATEGORY_KEYS,
    beforeScores,
  );
  profile.interventions = mergeInterventions(
    profile.interventions,
    Array.isArray(analysis.interventions) ? analysis.interventions : [],
  );

  const scoreDelta = {};
  for (const key of CATEGORY_KEYS) {
    scoreDelta[key] = profile.categoryScores[key] - beforeScores[key];
  }

  const assessment = analysis?.interactionAssessment || {};
  profile.interactionLog.push({
    id: crypto.randomUUID(),
    timestamp: now,
    summary:
      sanitizeText(assessment.summary, 320) ||
      "Interaction captured and profile updated.",
    progress: sanitizeText(assessment.progress, 20) || "stable",
    confidence: sanitizeText(assessment.confidence, 20) || "medium",
    assessmentMode: "analyzer",
    sessionPace: normalizeSessionPace(sessionPace),
    useCase: inferUseCaseFromText(userMessage, "General coaching"),
    userPrompt: sanitizeText(userMessage, 240),
    coachResponse: sanitizeText(assistantReply, 240),
    meetingNotesSnippet: sanitizeText(meetingNotes, 240),
    selfReportSnippet: sanitizeText(selfReport, 240),
    coacheeFeedback: null,
    categoryScores: { ...profile.categoryScores },
    scoreDelta,
    evidenceFlags: {
      meetingNotesProvided: Boolean(sanitizeText(meetingNotes, 8)),
      selfReportProvided: Boolean(sanitizeText(selfReport, 8)),
    },
  });

  const cutoff = retentionCutoffMs();
  profile.interactionLog = profile.interactionLog
    .filter((entry) => toMs(entry.timestamp) >= cutoff)
    .slice(-400);
  profile.interventions = profile.interventions.filter(
    (item) => toMs(item.updatedAt || item.createdAt) >= cutoff,
  );
  profile.feedbackLoop = normalizeFeedbackLoop({
    ...profile.feedbackLoop,
    lastAssessmentAt: now,
    lastAssessmentSource: "analyzer",
    lastInterventionReviewAt: now,
  });
  profile.feedbackLoop = recomputeFeedbackLoop(profile);
  profile.metrics = computeProfileMetrics(profile);
  profile.updatedAt = now;
  user.coachingProfile = normalizeCoachingProfile(profile);
}

function recordFallbackCoacheeAssessment({
  user,
  userMessage,
  assistantReply,
  meetingNotes,
  selfReport,
  sessionPace,
  reason,
}) {
  const profile = ensureUserCoachingProfile(user);
  const now = newIsoNow();
  const beforeScores = { ...profile.categoryScores };
  const scoreDelta = {};
  for (const key of CATEGORY_KEYS) {
    scoreDelta[key] = 0;
  }

  const summaryParts = [];
  const cleanReason = sanitizeText(reason, 180);
  const cleanUserMsg = sanitizeText(userMessage, 160);
  const cleanReply = sanitizeText(assistantReply, 160);
  if (cleanUserMsg) summaryParts.push(`User: ${cleanUserMsg}`);
  if (cleanReply) summaryParts.push(`Coach: ${cleanReply}`);
  if (!summaryParts.length) summaryParts.push("Interaction captured.");

  profile.interactionLog.push({
    id: crypto.randomUUID(),
    timestamp: now,
    summary: sanitizeText(
      `${summaryParts.join(" | ")}${cleanReason ? ` | Analysis pending: ${cleanReason}` : ""}`,
      320,
    ),
    progress: "stable",
    confidence: "low",
    assessmentMode: "fallback",
    sessionPace: normalizeSessionPace(sessionPace),
    useCase: inferUseCaseFromText(userMessage, "General coaching"),
    userPrompt: sanitizeText(userMessage, 240),
    coachResponse: sanitizeText(assistantReply, 240),
    meetingNotesSnippet: sanitizeText(meetingNotes, 240),
    selfReportSnippet: sanitizeText(selfReport, 240),
    coacheeFeedback: null,
    categoryScores: { ...beforeScores },
    scoreDelta,
    evidenceFlags: {
      meetingNotesProvided: Boolean(sanitizeText(meetingNotes, 8)),
      selfReportProvided: Boolean(sanitizeText(selfReport, 8)),
    },
  });

  const cutoff = retentionCutoffMs();
  profile.interactionLog = profile.interactionLog
    .filter((entry) => toMs(entry.timestamp) >= cutoff)
    .slice(-400);
  profile.interventions = profile.interventions.filter(
    (item) => toMs(item.updatedAt || item.createdAt) >= cutoff,
  );
  profile.feedbackLoop = normalizeFeedbackLoop({
    ...profile.feedbackLoop,
    lastAssessmentAt: now,
    lastAssessmentSource: "fallback",
  });
  profile.feedbackLoop = recomputeFeedbackLoop(profile);
  profile.metrics = computeProfileMetrics(profile);
  profile.updatedAt = now;
  user.coachingProfile = normalizeCoachingProfile(profile);
}

async function updateProfileAfterAssistantReply({
  apiKey,
  user,
  userMessage,
  assistantReply,
  meetingNotes,
  selfReport,
  sessionPace,
}) {
  try {
    await analyzeAndUpdateCoacheeProfile({
      apiKey,
      user,
      userMessage,
      assistantReply,
      meetingNotes,
      selfReport,
      sessionPace,
    });
    user.updatedAt = newIsoNow();
    await persistUsers();
  } catch (err) {
    console.error("Profile analysis failed:", err.message);
    try {
      recordFallbackCoacheeAssessment({
        user,
        userMessage,
        assistantReply,
        meetingNotes,
        selfReport,
        sessionPace,
        reason: err.message || "Analyzer unavailable.",
      });
      user.updatedAt = newIsoNow();
      await persistUsers();
    } catch (fallbackErr) {
      console.error("Fallback assessment failed:", fallbackErr.message);
    }
  }
}

async function handleChat(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }
  if (user.mustChangePassword) {
    sendJson(res, 403, {
      error: "Password change required before using coaching chat.",
    });
    return;
  }

  const key = getValidatedApiKey();
  if (!key.ok) {
    sendJson(res, 500, { error: key.error });
    return;
  }

  try {
    const body = await parseBody(req);
    const incoming = Array.isArray(body.messages) ? body.messages : [];
    const initiate = Boolean(body.initiate);
    const responseMode = sanitizeText(body.responseMode, 40).toLowerCase();
    const voiceFastMode = responseMode === "voice_fast";
    const sessionPace = String(body.sessionPace || "standard")
      .trim()
      .toLowerCase() === "busy"
      ? "busy"
      : "standard";
    const meetingNotes = sanitizeText(body.meetingNotes, 6000);
    const selfReport = sanitizeText(body.selfReport, 3000);
    const messages = incoming
      .filter((m) => m && (m.role === "user" || m.role === "assistant"))
      .map((m) => ({ role: m.role, content: String(m.content || "") }));
    const historyLimit = voiceFastMode
      ? Math.min(10, MAX_CHAT_HISTORY_MESSAGES)
      : MAX_CHAT_HISTORY_MESSAGES;
    const recentMessages = messages.slice(-historyLimit);
    const latestUserMessage =
      [...messages].reverse().find((m) => m.role === "user")?.content || "";
    const fastTurn = voiceFastMode || sessionPace === "busy";

    const payload = await createOpenAIChatCompletion(key.apiKey, {
      model: MODEL,
      temperature: fastTurn ? 0.25 : 0.4,
      max_tokens: fastTurn ? 170 : 300,
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "system", content: COACHING_CONSTITUTION_PROMPT },
        { role: "system", content: buildUserContext(user) },
        {
          role: "system",
          content: initiate
            ? "Initiate naturally in 1-3 short sentences. Start by asking the coachee what outcome/goal they want from this conversation."
            : "Continue naturally in short dialogue. If asked for a full framework cycle, deliver it through conversation across turns rather than list-heavy stages, unless explicitly asked for a structured summary.",
        },
        {
          role: "system",
          content:
            sessionPace === "busy"
              ? "Session pace is BUSY. Prioritize speed: keep each reply to 1-3 short sentences, lead with the highest-value recommendation or question first, and end with one concrete next action."
              : "Session pace is STANDARD. Keep replies concise but allow brief nuance when useful.",
        },
        {
          role: "system",
          content: voiceFastMode
            ? "Voice-fast mode: respond in 1-2 very short spoken sentences with one concrete next step or question."
            : "Standard mode: concise conversational response.",
        },
        {
          role: "system",
          content:
            meetingNotes || selfReport
              ? `Session evidence:\n- Meeting notes: ${
                  meetingNotes || "None"
                }\n- Coachee self-report: ${selfReport || "None"}`
              : "Session evidence: no additional notes provided.",
        },
        ...recentMessages,
      ],
    });

    const reply = payload?.choices?.[0]?.message?.content;
    if (!reply) {
      sendJson(res, 502, { error: "No assistant reply returned." });
      return;
    }

    sendJson(res, 200, { reply, usage: payload.usage || null });
    if (!initiate && latestUserMessage) {
      // Run profile analysis asynchronously so chat replies return immediately.
      void updateProfileAfterAssistantReply({
        apiKey: key.apiKey,
        user,
        userMessage: latestUserMessage,
        assistantReply: reply,
        meetingNotes,
        selfReport,
        sessionPace,
      });
    }
    return;
  } catch (err) {
    const status = Number(err?.status || 0);
    if (status >= 400 && status < 600) {
      sendJson(res, status, { error: err.message || "OpenAI API request failed." });
      return;
    }
    const cause = err?.cause?.message ? ` Cause: ${err.cause.message}` : "";
    sendJson(res, 502, {
      error: `Network error reaching OpenAI.${cause}`,
    });
  }
}

async function handleVoice(req, res) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    unauthorized(res);
    return;
  }
  if (user.mustChangePassword) {
    sendJson(res, 403, {
      error: "Password change required before using voice coaching.",
    });
    return;
  }

  const key = getValidatedApiKey();
  if (!key.ok) {
    sendJson(res, 500, { error: key.error });
    return;
  }

  try {
    const body = await parseBody(req);
    const text = sanitizeText(body.text, 3600);
    if (!text) {
      sendJson(res, 400, { error: "Provide text for speech synthesis." });
      return;
    }

    const voice = normalizeTtsVoice(body.voice, DEFAULT_TTS_VOICE);
    const format = normalizeTtsFormat(body.format, DEFAULT_TTS_FORMAT);
    const speedRaw = Number.parseFloat(body.speed);
    const speed = Number.isFinite(speedRaw)
      ? Math.max(0.7, Math.min(1.2, speedRaw))
      : 1;
    const instructions = sanitizeText(
      body.instructions,
      240,
    );

    const audioBuffer = await createOpenAISpeech(key.apiKey, {
      model: TTS_MODEL,
      voice,
      input: text,
      format,
      speed,
      instructions: instructions || undefined,
    });

    res.writeHead(200, {
      "Content-Type": getTtsContentType(format),
      "Content-Length": audioBuffer.length,
      "Cache-Control": "no-store",
    });
    res.end(audioBuffer);
  } catch (err) {
    const status = Number(err?.status || 0);
    if (status >= 400 && status < 600) {
      sendJson(res, status, { error: err.message || "TTS request failed." });
      return;
    }
    const cause = err?.cause?.message ? ` Cause: ${err.cause.message}` : "";
    sendJson(res, 502, {
      error: `Network error reaching OpenAI voice service.${cause}`,
    });
  }
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url || "/", `http://${req.headers.host || HOST}`);
    const pathname = url.pathname;

    if (req.method === "POST" && pathname === "/api/register") {
      await handleRegister(req, res);
      return;
    }
    if (req.method === "POST" && pathname === "/api/login") {
      await handleLogin(req, res);
      return;
    }
    if (req.method === "POST" && pathname === "/api/logout") {
      handleLogout(req, res);
      return;
    }
    if (req.method === "GET" && pathname === "/api/me") {
      handleMe(req, res);
      return;
    }
    if (req.method === "GET" && pathname === "/api/dashboard") {
      handleDashboard(req, res);
      return;
    }
    if (req.method === "GET" && pathname === "/api/admin/dashboard") {
      handleAdminDashboard(req, res);
      return;
    }
    if (req.method === "PUT" && pathname === "/api/admin/settings") {
      await handleAdminSettingsUpdate(req, res);
      return;
    }
    if (req.method === "POST" && pathname === "/api/admin/users") {
      await handleAdminCreateUser(req, res);
      return;
    }
    if (req.method === "GET" && pathname === "/api/personas") {
      handleStakeholderPersonas(req, res);
      return;
    }
    if (req.method === "GET" && pathname === "/api/playbooks") {
      handleDomainPlaybooks(req, res);
      return;
    }
    if (req.method === "GET" && pathname === "/api/practice-plans") {
      await handlePracticePlansList(req, res);
      return;
    }
    if (req.method === "POST" && pathname === "/api/practice-plans") {
      await handlePracticePlansCreate(req, res);
      return;
    }
    const practicePlanMatch = pathname.match(/^\/api\/practice-plans\/([^/]+)$/);
    if (req.method === "PUT" && practicePlanMatch) {
      await handlePracticePlanUpdate(
        req,
        res,
        decodeURIComponent(practicePlanMatch[1]),
      );
      return;
    }
    const adminActionMatch = pathname.match(/^\/api\/admin\/users\/([^/]+)\/action$/);
    if (req.method === "POST" && adminActionMatch) {
      await handleAdminUserAction(req, res, decodeURIComponent(adminActionMatch[1]));
      return;
    }
    if (req.method === "PUT" && pathname === "/api/profile") {
      await handleProfileUpdate(req, res);
      return;
    }
    if (req.method === "PUT" && pathname === "/api/password") {
      await handlePasswordChange(req, res);
      return;
    }
    if (req.method === "POST" && pathname === "/api/feedback") {
      await handleFeedback(req, res);
      return;
    }
    if (req.method === "GET" && pathname === "/api/nudges") {
      await handleNudgesList(req, res);
      return;
    }
    if (req.method === "POST" && pathname === "/api/nudges") {
      await handleNudgesCreate(req, res);
      return;
    }
    const nudgeMatch = pathname.match(/^\/api\/nudges\/([^/]+)$/);
    if (req.method === "PUT" && nudgeMatch) {
      await handleNudgeUpdate(req, res, decodeURIComponent(nudgeMatch[1]));
      return;
    }
    if (req.method === "POST" && pathname === "/api/chat") {
      await handleChat(req, res);
      return;
    }
    if (req.method === "POST" && pathname === "/api/voice") {
      await handleVoice(req, res);
      return;
    }
    if (req.method === "GET") {
      await serveStatic(pathname, res);
      return;
    }

    res.writeHead(405);
    res.end("Method not allowed");
  } catch (err) {
    sendJson(res, 500, { error: err.message || "Unexpected server error." });
  }
});

async function start() {
  await loadUsers();
  server.listen(PORT, HOST, () => {
    console.log(`Executive coach voice app running at http://${HOST}:${PORT}`);
    console.log(`Model: ${MODEL}`);
    console.log(`TTS model: ${TTS_MODEL} (${DEFAULT_TTS_VOICE}/${DEFAULT_TTS_FORMAT})`);
  });
}

start().catch((err) => {
  console.error("Failed to start server:", err.message);
  process.exit(1);
});
