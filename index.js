const { app } = require('@azure/functions');

// Load and register your existing v4 function (do not move it)
require('./src/functions/ProcessVttFile');

module.exports = app;

function safeJsonParse(str) {
  try { return JSON.parse(str); } catch { return null; }
}

function deriveKeyPointsFallbackFromText(text) {
  if (!text) return [];
  // Collect bullet-style lines
  const bullets = Array.from(new Set(
    text.split(/\r?\n+/)
        .filter(l => /^\s*[-•–]/.test(l))
        .map(l => l.replace(/^\s*[-•–]\s*/, "").trim())
        .filter(Boolean)
  ));
  if (bullets.length >= 3) return bullets.slice(0, 12);
  // Otherwise take first sentences
  return text
    .split(/(?<=[.!?])\s+/)
    .map(s => s.trim())
    .filter(Boolean)
    .slice(0, 8);
}
// ---------- End helpers ----------