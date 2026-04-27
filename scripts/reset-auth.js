const fs = require("fs");
const path = require("path");

const rootDir = path.resolve(__dirname, "..");
const dataDir = path.join(rootDir, ".data");
const usersFile = path.join(dataDir, "users.json");
const sessionsDir = path.join(dataDir, "sessions");

removeIfExists(usersFile);
removeIfExists(sessionsDir);
fs.mkdirSync(sessionsDir, { recursive: true });

process.stdout.write("Authentication store reset. Run `npm start` to reseed the admin account.\n");

function removeIfExists(targetPath) {
  if (!fs.existsSync(targetPath)) return;
  fs.rmSync(targetPath, { recursive: true, force: true });
}
