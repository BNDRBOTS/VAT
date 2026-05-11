import { createReadStream, existsSync, statSync } from "node:fs";
import { extname, join, normalize } from "node:path";
import http from "node:http";

const DIST_DIR = join(process.cwd(), "dist");
const PORT = Number(process.env.PORT || 3000);
const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".ico": "image/x-icon",
  ".woff": "font/woff",
  ".woff2": "font/woff2"
};
function safePath(urlPath) {
  const pathname = decodeURIComponent((urlPath || "/").split("?")[0]);
  const normalizedPath = normalize(pathname).replace(/^(\.\.[/\\])+/, "");
  return join(DIST_DIR, normalizedPath);
}
function sendFile(res, filePath) {
  const ext = extname(filePath).toLowerCase();
  const stat = statSync(filePath);
  res.writeHead(200, {
    "Content-Type": MIME[ext] || "application/octet-stream",
    "Content-Length": stat.size,
    "Cache-Control": ext === ".html" ? "no-cache" : "public, max-age=31536000, immutable"
  });
  createReadStream(filePath).pipe(res);
}
const server = http.createServer((req, res) => {
  const requested = safePath(req.url || "/");
  const indexPath = join(DIST_DIR, "index.html");
  if (existsSync(requested) && statSync(requested).isFile()) return sendFile(res, requested);
  const htmlCandidate = requested.endsWith(".html") ? requested : requested + ".html";
  if (existsSync(htmlCandidate) && statSync(htmlCandidate).isFile()) return sendFile(res, htmlCandidate);
  if (existsSync(indexPath)) return sendFile(res, indexPath);
  res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
  res.end("Build output missing. Run npm run build first.");
});
server.listen(PORT, "0.0.0.0", () => console.log(`VAT forensic app listening on ${PORT}`));