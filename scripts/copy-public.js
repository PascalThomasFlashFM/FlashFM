// Copie src/web/public -> dist/web/public apres compilation TypeScript.
// En JS pur (pas de commande shell) pour fonctionner a l'identique sur
// Windows, macOS et Linux.
const fs = require("fs");
const path = require("path");

const src = path.join(__dirname, "..", "src", "web", "public");
const dest = path.join(__dirname, "..", "dist", "web", "public");

fs.mkdirSync(dest, { recursive: true });
fs.cpSync(src, dest, { recursive: true });
