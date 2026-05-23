const selfsigned = require("selfsigned");
const fs   = require("fs");
const path = require("path");

const certsDir = path.join(__dirname, "..", "certs");
if (!fs.existsSync(certsDir)) fs.mkdirSync(certsDir);

const attrs = [{ name: "commonName", value: "localhost" }];
const opts  = {
  keySize:  2048,
  days:     365,
  algorithm: "sha256",
  extensions: [
    { name: "subjectAltName", altNames: [{ type: 2, value: "localhost" }, { type: 7, ip: "127.0.0.1" }] },
  ],
};

console.log("Generating self-signed certificate...");
const pems = selfsigned.generate(attrs, opts);

fs.writeFileSync(path.join(certsDir, "server.key"),  pems.private);
fs.writeFileSync(path.join(certsDir, "server.crt"),  pems.cert);

console.log("Done! Certs written to certs/server.key and certs/server.crt");
console.log("\nNext: open https://localhost:3000 in your browser and trust the certificate.");
