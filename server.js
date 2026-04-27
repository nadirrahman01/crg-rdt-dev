const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

require("dotenv").config();

const express = require("express");
const helmet = require("helmet");
const session = require("express-session");
const FileStore = require("session-file-store")(session);

const ROOT_DIR = __dirname;
const DATA_DIR = path.join(ROOT_DIR, ".data");
const SESSION_DIR = path.join(DATA_DIR, "sessions");
const USERS_FILE = path.join(DATA_DIR, "users.json");
const SYSTEM_FILE = path.join(DATA_DIR, "system.json");

const PORT = Number(process.env.PORT || 3000);
const ADMIN_EMAIL = String(process.env.ADMIN_EMAIL || "nadir@cordobarg.com").trim().toLowerCase();
const ADMIN_PASSWORD = String(process.env.ADMIN_PASSWORD || "nadir900");

ensureDirectory(DATA_DIR);
ensureDirectory(SESSION_DIR);

const app = express();
const systemState = loadJsonFile(SYSTEM_FILE, {});
const sessionSecret = process.env.SESSION_SECRET || systemState.sessionSecret || generateSessionSecret();

if (!systemState.sessionSecret && !process.env.SESSION_SECRET) {
  saveJsonFile(SYSTEM_FILE, {
    ...systemState,
    sessionSecret
  });
}

app.disable("x-powered-by");
app.set("trust proxy", 1);
app.use(helmet({
  crossOriginEmbedderPolicy: false,
  contentSecurityPolicy: false
}));
app.use(express.json({ limit: "1mb" }));
app.use(express.urlencoded({ extended: false }));
app.use(session({
  name: "rdt.sid",
  secret: sessionSecret,
  resave: false,
  saveUninitialized: false,
  store: new FileStore({
    path: SESSION_DIR,
    retries: 0,
    ttl: 60 * 60 * 24 * 30
  }),
  cookie: {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
    maxAge: 1000 * 60 * 60 * 8
  }
}));

app.use("/assets", express.static(path.join(ROOT_DIR, "assets")));

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, service: "rdt-auth" });
});

app.get("/api/auth/session", async (req, res) => {
  if (!req.session.user) {
    return res.json({ authenticated: false });
  }

  const user = await getUserById(req.session.user.id);
  if (!user) {
    req.session.destroy(() => undefined);
    return res.json({ authenticated: false });
  }

  res.json({
    authenticated: true,
    user: buildSessionUser(user),
    security: {
      passwordChangedAt: user.passwordChangedAt || user.updatedAt || user.createdAt,
      twoFactorEnabled: Boolean(user.twoFactorEnabled),
      activeSessions: countSessionsForUser(user.id)
    }
  });
});

app.post("/api/auth/login", async (req, res) => {
  const email = String(req.body?.email || "").trim().toLowerCase();
  const password = String(req.body?.password || "");
  const remember = Boolean(req.body?.remember);

  if (!email || !password) {
    return res.status(400).json({ error: "Enter your work email and password." });
  }

  const user = await getUserByEmail(email);
  if (!user || !(await verifyPassword(password, user.passwordHash))) {
    return res.status(401).json({ error: "Invalid email or password." });
  }

  req.session.regenerate((sessionError) => {
    if (sessionError) {
      return res.status(500).json({ error: "Unable to start a secure session." });
    }

    req.session.user = buildSessionUser(user);
    req.session.cookie.maxAge = remember
      ? 1000 * 60 * 60 * 24 * 30
      : 1000 * 60 * 60 * 8;

    req.session.save((saveError) => {
      if (saveError) {
        return res.status(500).json({ error: "Unable to save the current session." });
      }

      return res.json({
        authenticated: true,
        user: buildSessionUser(user),
        security: {
          passwordChangedAt: user.passwordChangedAt || user.updatedAt || user.createdAt,
          twoFactorEnabled: Boolean(user.twoFactorEnabled),
          activeSessions: countSessionsForUser(user.id)
        }
      });
    });
  });
});

app.post("/api/auth/logout", requireAuth, (req, res) => {
  req.session.destroy(() => {
    res.clearCookie("rdt.sid");
    res.json({ ok: true });
  });
});

app.post("/api/auth/change-password", requireAuth, async (req, res) => {
  const currentPassword = String(req.body?.currentPassword || "");
  const newPassword = String(req.body?.newPassword || "");

  if (!currentPassword || !newPassword || newPassword.length < 8) {
    return res.status(400).json({ error: "Enter your current password and a stronger new password." });
  }

  const users = loadUsers();
  const userIndex = users.findIndex((entry) => entry.id === req.session.user.id);
  if (userIndex < 0) {
    return res.status(404).json({ error: "User account not found." });
  }

  const user = users[userIndex];
  if (!(await verifyPassword(currentPassword, user.passwordHash))) {
    return res.status(401).json({ error: "Current password is incorrect." });
  }

  users[userIndex] = {
    ...user,
    passwordHash: await hashPassword(newPassword),
    passwordChangedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  saveUsers(users);

  res.json({
    ok: true,
    passwordChangedAt: users[userIndex].passwordChangedAt
  });
});

app.get("/api/settings", requireAuth, async (req, res) => {
  const user = await getUserById(req.session.user.id);
  if (!user) return res.status(404).json({ error: "User account not found." });

  res.json(buildSettingsPayload(user));
});

app.post("/api/settings", requireAuth, async (req, res) => {
  const users = loadUsers();
  const userIndex = users.findIndex((entry) => entry.id === req.session.user.id);
  if (userIndex < 0) {
    return res.status(404).json({ error: "User account not found." });
  }

  const payload = req.body || {};
  const current = users[userIndex];
  users[userIndex] = {
    ...current,
    fullName: safeString(payload.profile?.fullName, current.fullName),
    phoneNumber: safeString(payload.profile?.phoneNumber, current.phoneNumber),
    jobTitle: safeString(payload.profile?.jobTitle, current.jobTitle),
    defaultNoteType: safeString(payload.defaults?.defaultNoteType, current.defaultNoteType),
    defaultRegion: safeString(payload.defaults?.defaultRegion, current.defaultRegion),
    notifications: {
      publicationReminders: Boolean(payload.notifications?.publicationReminders),
      draftActivity: Boolean(payload.notifications?.draftActivity),
      validationAlerts: Boolean(payload.notifications?.validationAlerts),
      systemUpdates: Boolean(payload.notifications?.systemUpdates)
    },
    updatedAt: new Date().toISOString()
  };

  saveUsers(users);
  res.json(buildSettingsPayload(users[userIndex]));
});

app.get("/", (_req, res) => {
  res.sendFile(path.join(ROOT_DIR, "index.html"));
});

app.get("*", (req, res, next) => {
  if (req.path.startsWith("/api/")) return next();
  res.sendFile(path.join(ROOT_DIR, "index.html"));
});

bootstrap().then(() => {
  app.listen(PORT, () => {
    process.stdout.write(`RDT server listening on http://localhost:${PORT}\n`);
  });
}).catch((error) => {
  process.stderr.write(`${error.stack || error.message}\n`);
  process.exit(1);
});

async function bootstrap() {
  const users = loadUsers();
  const existingIndex = users.findIndex((entry) => entry.email === ADMIN_EMAIL);
  const passwordHash = await hashPassword(ADMIN_PASSWORD);
  const now = new Date().toISOString();
  const adminRecord = {
    id: existingIndex >= 0 ? users[existingIndex].id : crypto.randomUUID(),
    email: ADMIN_EMAIL,
    fullName: existingIndex >= 0 ? users[existingIndex].fullName : "Nadir Rahman",
    phoneNumber: existingIndex >= 0 ? users[existingIndex].phoneNumber : "+44 20 7946 0118",
    jobTitle: existingIndex >= 0 ? users[existingIndex].jobTitle : "Administrator / Research Production",
    role: "admin",
    passwordHash,
    passwordChangedAt: now,
    twoFactorEnabled: true,
    defaultNoteType: existingIndex >= 0 ? users[existingIndex].defaultNoteType : "Macro / Sovereign Outlook",
    defaultRegion: existingIndex >= 0 ? users[existingIndex].defaultRegion : "Emerging Markets",
    notifications: {
      publicationReminders: true,
      draftActivity: true,
      validationAlerts: true,
      systemUpdates: true,
      ...(existingIndex >= 0 ? users[existingIndex].notifications : {})
    },
    createdAt: existingIndex >= 0 ? users[existingIndex].createdAt : now,
    updatedAt: now,
    seedManaged: true
  };

  if (existingIndex >= 0) {
    users[existingIndex] = adminRecord;
  } else {
    users.push(adminRecord);
  }

  saveUsers(users);
}

function requireAuth(req, res, next) {
  if (!req.session.user?.id) {
    return res.status(401).json({ error: "Authentication required." });
  }
  return next();
}

function buildSessionUser(user) {
  return {
    id: user.id,
    email: user.email,
    fullName: user.fullName,
    role: user.role,
    initials: initialsForName(user.fullName)
  };
}

function buildSettingsPayload(user) {
  return {
    profile: {
      fullName: user.fullName || "",
      email: user.email || "",
      phoneNumber: user.phoneNumber || "",
      jobTitle: user.jobTitle || ""
    },
    defaults: {
      defaultNoteType: user.defaultNoteType || "Macro / Sovereign Outlook",
      defaultRegion: user.defaultRegion || "Emerging Markets"
    },
    notifications: {
      publicationReminders: Boolean(user.notifications?.publicationReminders),
      draftActivity: Boolean(user.notifications?.draftActivity),
      validationAlerts: Boolean(user.notifications?.validationAlerts),
      systemUpdates: Boolean(user.notifications?.systemUpdates)
    },
    security: {
      passwordChangedAt: user.passwordChangedAt || user.updatedAt || user.createdAt,
      twoFactorEnabled: Boolean(user.twoFactorEnabled),
      activeSessions: countSessionsForUser(user.id)
    }
  };
}

async function getUserByEmail(email) {
  return loadUsers().find((entry) => entry.email === email) || null;
}

async function getUserById(id) {
  return loadUsers().find((entry) => entry.id === id) || null;
}

function loadUsers() {
  return loadJsonFile(USERS_FILE, []);
}

function saveUsers(users) {
  saveJsonFile(USERS_FILE, users);
}

function loadJsonFile(filePath, fallback) {
  try {
    if (!fs.existsSync(filePath)) {
      saveJsonFile(filePath, fallback);
      return fallback;
    }
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch (_error) {
    return fallback;
  }
}

function saveJsonFile(filePath, value) {
  ensureDirectory(path.dirname(filePath));
  fs.writeFileSync(filePath, JSON.stringify(value, null, 2));
}

function ensureDirectory(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function generateSessionSecret() {
  return crypto.randomBytes(32).toString("hex");
}

async function hashPassword(password) {
  const salt = crypto.randomBytes(16).toString("hex");
  const derivedKey = await scrypt(password, salt, 64);
  return {
    algorithm: "scrypt",
    salt,
    iterations: 16384,
    keyLength: 64,
    digest: derivedKey.toString("hex")
  };
}

async function verifyPassword(password, passwordHash) {
  if (!passwordHash || passwordHash.algorithm !== "scrypt" || !passwordHash.salt || !passwordHash.digest) {
    return false;
  }

  const derivedKey = await scrypt(password, passwordHash.salt, passwordHash.keyLength || 64);
  const digestBuffer = Buffer.from(passwordHash.digest, "hex");
  if (digestBuffer.length !== derivedKey.length) return false;
  return crypto.timingSafeEqual(derivedKey, digestBuffer);
}

function scrypt(password, salt, keyLength) {
  return new Promise((resolve, reject) => {
    crypto.scrypt(password, salt, keyLength, { N: 16384, r: 8, p: 1 }, (error, derivedKey) => {
      if (error) reject(error);
      else resolve(derivedKey);
    });
  });
}

function initialsForName(name) {
  return String(name || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part.charAt(0).toUpperCase())
    .join("") || "CR";
}

function safeString(value, fallback = "") {
  const next = String(value ?? "").trim();
  return next || fallback;
}

function countSessionsForUser(userId) {
  try {
    const files = fs.readdirSync(SESSION_DIR).filter((file) => file.endsWith(".json"));
    return files.reduce((count, fileName) => {
      const sessionJson = JSON.parse(fs.readFileSync(path.join(SESSION_DIR, fileName), "utf8"));
      return sessionJson?.user?.id === userId ? count + 1 : count;
    }, 0) || 1;
  } catch (_error) {
    return 1;
  }
}
