const https   = require("https");
const fs      = require("fs");
const path    = require("path");
const devCerts = require("office-addin-dev-certs");

const PORT = 3000;

const MIME = {
  ".html": "text/html",
  ".css":  "text/css",
  ".js":   "application/javascript",
  ".png":  "image/png",
  ".xml":  "application/xml",
  ".json": "application/json",
};

async function start() {
  let sslOptions;
  try {
    sslOptions = await devCerts.getHttpsServerOptions();
  } catch {
    console.error("\nSSL certificates not found. Run: npm run setup-certs\n");
    process.exit(1);
  }

  const server = https.createServer(sslOptions, (req, res) => {
    let urlPath = req.url.split("?")[0];
    if (urlPath === "/") urlPath = "/taskpane.html";

    const filePath = path.join(__dirname, urlPath);

    // Prevent path traversal
    if (!filePath.startsWith(__dirname)) {
      res.writeHead(403);
      res.end("Forbidden");
      return;
    }

    fs.readFile(filePath, (err, data) => {
      if (err) {
        res.writeHead(404, { "Content-Type": "text/plain" });
        res.end("Not found");
        return;
      }
      const ext  = path.extname(filePath).toLowerCase();
      const mime = MIME[ext] || "application/octet-stream";
      res.writeHead(200, {
        "Content-Type": mime,
        "Access-Control-Allow-Origin": "*",
      });
      res.end(data);
    });
  });

  server.listen(PORT, () => {
    console.log(`\nCalendar Event Extractor running at https://localhost:${PORT}\n`);
    console.log("  1. Make sure https://localhost:3000 loads without cert warnings in Edge");
    console.log("  2. Sideload manifest.xml in Outlook Web (Get Add-ins > My add-ins > Add from file)");
  });
}

start();
