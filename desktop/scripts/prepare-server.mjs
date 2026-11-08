// Puts together what "On this computer" runs: the standalone Dossier server, its
// database setup, and a Node binary, for one platform, in build-server/. Every
// download is checked against a hash pinned here. Run after the app's own build:
//
//   node scripts/prepare-server.mjs --next ../.next --seed ../.next/desktop-seed.mjs --platform win32

import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { chmodSync, cpSync, existsSync, lstatSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { parseArgs } from "node:util";

const NODE_VERSION = "24.22.2";
const NODE_ABI = "137";
const BETTER_SQLITE3_VERSION = "12.11.1";
const SHARP_VERSION = "0.35.5";

const DOWNLOADS = {
  linux: {
    node: { url: `https://nodejs.org/dist/v${NODE_VERSION}/node-v${NODE_VERSION}-linux-x64.tar.gz`, sha256: "361718c444e49af2ca5a110c83ba6a8901143f23ab285ebbef0341f391f482a9" },
    sqlite: { sha256: "99c43785639d5d3690c396ba245ee680ac8b469a46b19233a3546f0eb7f8e312" },
    sharp: [
      { name: "sharp-linux-x64", version: SHARP_VERSION, sha256: "556787285d3244c07686a355de5b5a4c4d0019a12785bf24764e756a96ea3334" },
      { name: "sharp-libvips-linux-x64", version: "1.3.4", sha256: "e371cb713fc8822c96a0d3fb60a3d092799d0aefab6d7f1fb3b72d8d75f1d109" },
    ],
  },
  win32: {
    node: { url: `https://nodejs.org/dist/v${NODE_VERSION}/win-x64/node.exe`, sha256: "1c331340fa6cac5e4294725959070a143fa0507195a78ad77eed7943f719aac9" },
    sqlite: { sha256: "4ee5e653174d6ddd301605d351798cdae2613da06c4f37ecdce263021fcf1255" },
    // The Windows package includes libvips beside the addon.
    sharp: [
      { name: "sharp-win32-x64", version: SHARP_VERSION, sha256: "16a804181ddd258ca6fe7079b2faa0c0356654c9db3d93ab0a80ec5e5161c3d4" },
    ],
  },
};

const desktop = path.join(import.meta.dirname, "..");
const { values } = parseArgs({
  options: {
    next: { type: "string", default: path.join(desktop, "..", ".next") },
    seed: { type: "string", default: path.join(desktop, "..", ".next", "desktop-seed.mjs") },
    platform: { type: "string", default: process.platform },
  },
});
const downloads = DOWNLOADS[values.platform];
if (!downloads) throw new Error(`There's no desktop app for ${values.platform}.`);

async function download(url, sha256) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`${url} answered ${response.status}.`);
  const bytes = Buffer.from(await response.arrayBuffer());
  const actual = createHash("sha256").update(bytes).digest("hex");
  if (actual !== sha256) throw new Error(`${url} isn't the file expected: its hash is ${actual}.`);
  return bytes;
}

function run(command, args) {
  const result = spawnSync(command, args, { stdio: "inherit" });
  if (result.status !== 0) throw new Error(`${command} ${args.join(" ")} failed.`);
}

const out = path.join(desktop, "build-server");
const server = path.join(out, "server");
const scratch = mkdtempSync(path.join(tmpdir(), "dossier-prepare-"));
rmSync(out, { recursive: true, force: true });

/**
 * Replaces every link under a folder with what it points at: an installer can't
 * carry links, and these point at this machine. Node finds a package's
 * dependencies from where it really is, so a package nested in another's
 * node_modules needs the ones next to it brought along, or a copy can't find them.
 */
function replaceLinks(folder, topLevel) {
  for (const entry of readdirSync(folder)) {
    const entryPath = path.join(folder, entry);
    const stat = lstatSync(entryPath);
    if (stat.isSymbolicLink()) {
      const target = realpathSync(entryPath);
      rmSync(entryPath);
      cpSync(target, entryPath, { recursive: true, dereference: true });
      const parent = path.dirname(target);
      if (path.basename(parent) !== "node_modules" || parent === topLevel) continue;
      for (const sibling of readdirSync(parent)) {
        if (sibling !== path.basename(target)) cpSync(path.join(parent, sibling), path.join(entryPath, "node_modules", sibling), { recursive: true, dereference: true });
      }
    } else if (stat.isDirectory()) {
      replaceLinks(entryPath, topLevel);
    }
  }
}

/** Every file under a folder, for checking what's about to ship. */
function* filesUnder(folder) {
  for (const entry of readdirSync(folder)) {
    const entryPath = path.join(folder, entry);
    if (lstatSync(entryPath).isDirectory()) yield* filesUnder(entryPath);
    else yield entryPath;
  }
}

// The server. Next.js links its native packages from .next/node_modules.
cpSync(path.join(values.next, "standalone"), server, { recursive: true, verbatimSymlinks: true });
replaceLinks(path.join(server, ".next", "node_modules"), realpathSync(path.join(server, "node_modules")));
cpSync(path.join(values.next, "static"), path.join(server, ".next", "static"), { recursive: true });
cpSync(path.join(desktop, "..", "public"), path.join(server, "public"), { recursive: true });

// Only what the server needs, as a list of what to keep. The standalone folder
// also holds whatever Next.js traced from the project, and it copies .env into
// it: on a developer's machine, that's their real secrets.
const KEEP = new Set([".next", "node_modules", "package.json", "public", "server.js", "src"]);
for (const entry of readdirSync(server)) {
  if (!KEEP.has(entry)) rmSync(path.join(server, entry), { recursive: true, force: true });
}
// Only the Word runtime assets belong under src, not arbitrary traced sources.
const fonts = path.join(server, "src", "fonts", "word");
if (!existsSync(path.join(fonts, "manifest.json"))) throw new Error("The standalone build is missing Word fonts; rebuild with outputFileTracingIncludes.");
for (const entry of readdirSync(path.join(server, "src"))) {
  if (entry !== "fonts") rmSync(path.join(server, "src", entry), { recursive: true, force: true });
}
for (const entry of readdirSync(path.dirname(fonts))) {
  if (entry !== "word") rmSync(path.join(path.dirname(fonts), entry), { recursive: true, force: true });
}
const manifest = JSON.parse(readFileSync(path.join(fonts, "manifest.json"), "utf8"));
for (const face of manifest.faces) {
  const file = path.join(fonts, path.relative("word", face.file));
  if (!existsSync(file) || createHash("sha256").update(readFileSync(file)).digest("hex") !== face.sha256) {
    throw new Error(`The standalone build has a missing or stale Word font: ${face.file}.`);
  }
}

// Word photo masks use sharp even though next/image is disabled. Keep its JS
// and replace the builder's native addons with verified target-platform ones.
// Next may trace multiple copies at hashed paths, each with its own dependencies.
const sharpFolders = [...filesUnder(server)]
  .filter((file) => path.basename(file) === "package.json" && path.basename(path.dirname(file)).startsWith("sharp"))
  .filter((file) => JSON.parse(readFileSync(file, "utf8")).name === "sharp")
  .map((file) => path.dirname(file));
if (!sharpFolders.length) throw new Error("The standalone build is missing sharp, required for Word photo masks.");
function removeSharpNativePackages(scope) {
  if (!existsSync(scope)) return;
  for (const entry of readdirSync(scope)) {
    // @img/colour is portable JS required by sharp, not a native addon.
    if (entry.startsWith("sharp-")) rmSync(path.join(scope, entry), { recursive: true, force: true });
  }
}
for (const folder of sharpFolders) {
  const version = JSON.parse(readFileSync(path.join(folder, "package.json"), "utf8")).version;
  if (version !== SHARP_VERSION) throw new Error(`The build has sharp ${version}; this script pins ${SHARP_VERSION}.`);
  removeSharpNativePackages(path.join(folder, "node_modules", "@img"));
}
removeSharpNativePackages(path.join(server, "node_modules", "@img"));
for (const asset of downloads.sharp) {
  const archive = path.join(scratch, `${asset.name}.tgz`);
  const unpacked = path.join(scratch, asset.name);
  mkdirSync(unpacked, { recursive: true });
  writeFileSync(archive, await download(`https://registry.npmjs.org/@img/${asset.name}/-/${asset.name}-${asset.version}.tgz`, asset.sha256));
  run("tar", ["-xzf", archive, "-C", unpacked, "--strip-components=1"]);
  const pkg = JSON.parse(readFileSync(path.join(unpacked, "package.json"), "utf8"));
  if (pkg.name !== `@img/${asset.name}` || pkg.version !== asset.version) throw new Error(`Unexpected package in ${asset.name}.`);
  for (const folder of sharpFolders) {
    cpSync(unpacked, path.join(folder, "node_modules", "@img", asset.name), { recursive: true });
  }
}

// The SQLite driver, for this platform.
const sqliteAsset = `better-sqlite3-v${BETTER_SQLITE3_VERSION}-node-v${NODE_ABI}-${values.platform}-x64.tar.gz`;
const sqliteArchive = path.join(scratch, sqliteAsset);
writeFileSync(sqliteArchive, await download(`https://github.com/WiseLibs/better-sqlite3/releases/download/v${BETTER_SQLITE3_VERSION}/${sqliteAsset}`, downloads.sqlite.sha256));
run("tar", ["-xzf", sqliteArchive, "-C", scratch]);
const sqliteBinary = path.join(scratch, "build", "Release", "better_sqlite3.node");
const nextModules = path.join(server, ".next", "node_modules");
const sqliteFolders = [
  path.join(server, "node_modules", "better-sqlite3"),
  ...(existsSync(nextModules) ? readdirSync(nextModules).filter((name) => name.startsWith("better-sqlite3-")).map((name) => path.join(nextModules, name)) : []),
];
for (const folder of sqliteFolders) {
  const version = JSON.parse(readFileSync(path.join(folder, "package.json"), "utf8")).version;
  if (version !== BETTER_SQLITE3_VERSION) throw new Error(`The build has better-sqlite3 ${version}; this script pins ${BETTER_SQLITE3_VERSION}.`);
  cpSync(sqliteBinary, path.join(folder, "build", "Release", "better_sqlite3.node"));
}

// What sets the database up and starts the server.
cpSync(path.join(desktop, "..", "prisma", "sqlite", "migrations"), path.join(server, "migrations"), { recursive: true });
cpSync(values.seed, path.join(server, "seed.mjs"));
cpSync(path.join(desktop, "server-tools"), server, { recursive: true });

// Node.
const nodeFolder = path.join(out, "node");
mkdirSync(nodeFolder, { recursive: true });
if (values.platform === "win32") {
  writeFileSync(path.join(nodeFolder, "node.exe"), await download(downloads.node.url, downloads.node.sha256));
} else {
  const archive = path.join(scratch, "node.tar.gz");
  writeFileSync(archive, await download(downloads.node.url, downloads.node.sha256));
  run("tar", ["-xzf", archive, "-C", scratch, "--strip-components=2", `node-v${NODE_VERSION}-linux-x64/bin/node`]);
  cpSync(path.join(scratch, "node"), path.join(nodeFolder, "node"));
  chmodSync(path.join(nodeFolder, "node"), 0o755);
}

rmSync(scratch, { recursive: true, force: true });

// The last check before this can reach anyone.
for (const file of filesUnder(out)) {
  if (/^\.env/.test(path.basename(file))) throw new Error(`${file} must never ship: it can hold secrets.`);
  if (lstatSync(file).isSymbolicLink()) throw new Error(`${file} is a link, which an installer can't carry.`);
}
console.log(`Ready for ${values.platform}: ${out}`);
