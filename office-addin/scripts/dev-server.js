const fs = require("fs");
const http = require("http");
const https = require("https");
const path = require("path");
const {URL} = require("url");
const {getHttpsServerOptions} = require("office-addin-dev-certs/lib/httpsServerOptions");

const root = path.resolve(__dirname, "..");
const port = Number(process.env.PORT || "3000");
const host = process.env.HOST || "localhost";
const bridgeTarget = process.env.WORD_AI_BRIDGE_URL || "http://127.0.0.1:8765";
const useHttp = process.env.WORD_AI_ADDIN_HTTP === "1";

const contentTypes = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
};

function send(res, status, headers, body) {
  res.writeHead(status, headers);
  res.end(body);
}

function routeFile(urlPath) {
  if (urlPath === "/" || urlPath === "/taskpane.html") {
    return path.join(root, "src", "taskpane.html");
  }
  if (urlPath === "/taskpane.js") {
    return path.join(root, "dist", "taskpane.js");
  }
  if (urlPath.startsWith("/assets/")) {
    return path.join(root, urlPath.slice(1));
  }
  return null;
}

function serveStatic(req, res) {
  const parsed = new URL(req.url, `${useHttp ? "http" : "https"}://${req.headers.host || "localhost"}`);
  const filePath = routeFile(parsed.pathname);
  if (!filePath) {
    send(res, 404, {"Content-Type": "text/plain; charset=utf-8"}, "not found");
    return;
  }
  fs.readFile(filePath, (err, data) => {
    if (err) {
      send(res, 404, {"Content-Type": "text/plain; charset=utf-8"}, "not found");
      return;
    }
    send(res, 200, {
      "Content-Type": contentTypes[path.extname(filePath)] || "application/octet-stream",
      "Cache-Control": "no-store",
    }, data);
  });
}

function proxyBridge(req, res) {
  const target = new URL(req.url.replace(/^\/bridge/, "") || "/", bridgeTarget);
  const client = target.protocol === "https:" ? https : http;
  const headers = {...req.headers, host: target.host};
  delete headers.connection;
  const proxyReq = client.request({
    protocol: target.protocol,
    hostname: target.hostname,
    port: target.port,
    method: req.method,
    path: `${target.pathname}${target.search}`,
    headers,
  }, (proxyRes) => {
    res.writeHead(proxyRes.statusCode || 502, proxyRes.headers);
    proxyRes.pipe(res);
  });
  proxyReq.on("error", (err) => {
    send(res, 502, {"Content-Type": "application/json; charset=utf-8"}, JSON.stringify({ok: false, error: err.message}));
  });
  req.pipe(proxyReq);
}

async function main() {
  const handler = (req, res) => {
    if ((req.url || "").startsWith("/bridge")) {
      proxyBridge(req, res);
      return;
    }
    serveStatic(req, res);
  };

  const server = useHttp
    ? http.createServer(handler)
    : https.createServer(await getHttpsServerOptions(30, ["localhost", "127.0.0.1"]), handler);

  server.listen(port, host, () => {
    const scheme = useHttp ? "http" : "https";
    console.log(`Word AI add-in served at ${scheme}://${host}:${port}/taskpane.html`);
    console.log(`Bridge proxy target: ${bridgeTarget}`);
  });
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
