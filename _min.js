console.log("electron version:", process.versions.electron);
console.log("type:", process.type);
try { const m = require("electron/main"); console.log("electron/main:", typeof m); } catch(e) { console.log("electron/main err:", e.message); }
try { const m = process._linkedBinding("electron_browser_app"); console.log("linked:", typeof m); } catch(e) { console.log("linked err:", e.message); }
const Module = require("module");
const orig = Module._resolveFilename;
console.log("Module._cache keys with electron:", Object.keys(Module._cache).filter(k => k.includes("electron")).slice(0,5));
process.exit(0);
