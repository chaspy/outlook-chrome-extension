#!/usr/bin/env node

/* eslint-env node */

const fs = require("node:fs");
const { execSync } = require("node:child_process");

const [requireBumpFlag] = process.argv.slice(2);
const requireBump = requireBumpFlag === "--require-bump";

const VERSION_FILE = "VERSION";
const VERSION_PATTERN = /^v?(\d+)\.(\d+)\.(\d+)$/;

const readJson = (path) => JSON.parse(fs.readFileSync(path, "utf8"));
const readVersionFile = () => fs.readFileSync(VERSION_FILE, "utf8").trim();

const parseSemver = (value, sourceName) => {
  const match = VERSION_PATTERN.exec(value);
  if (!match) {
    throw new Error(
      `${sourceName} version "${value}" is invalid. Use format "X.Y.Z" or "vX.Y.Z".`
    );
  }

  return match.slice(1).map((part) => Number(part));
};

const compareSemver = (left, right) => {
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] > right[index]) return 1;
    if (left[index] < right[index]) return -1;
  }
  return 0;
};

const readLatestTag = () => {
  const raw = execSync("git tag --list 'v*' --sort=-v:refname | head -n 1", {
    encoding: "utf8"
  }).trim();
  return raw || null;
};

try {
  const canonicalVersion = readVersionFile();
  if (!canonicalVersion) {
    throw new Error(`${VERSION_FILE} is empty.`);
  }
  parseSemver(canonicalVersion, VERSION_FILE);

  const manifest = readJson("manifest.json");
  const pkg = readJson("package.json");
  const lock = readJson("package-lock.json");

  const manifestVersion = `${manifest.version ?? ""}`.trim();
  const packageVersion = `${pkg.version ?? ""}`.trim();
  const lockVersion = `${lock.version ?? ""}`.trim();
  const lockRootVersion = `${lock.packages?.[""]?.version ?? ""}`.trim();

  if (!manifestVersion) {
    throw new Error("manifest.json version is empty.");
  }
  if (!packageVersion) {
    throw new Error("package.json version is empty.");
  }
  if (!lockVersion) {
    throw new Error("package-lock.json version is empty.");
  }
  if (!lockRootVersion) {
    throw new Error("package-lock.json packages[\"\"].version is empty.");
  }

  const versions = [
    ["manifest.json", manifestVersion],
    ["package.json", packageVersion],
    ["package-lock.json", lockVersion],
    ["package-lock.json packages[\"\"]", lockRootVersion]
  ];

  for (const [sourceName, version] of versions) {
    if (version !== canonicalVersion) {
      throw new Error(
        `${sourceName} version=${version} does not match ${VERSION_FILE}=${canonicalVersion}`
      );
    }
  }

  if (manifestVersion !== packageVersion || packageVersion !== lockVersion || lockVersion !== lockRootVersion) {
    throw new Error(
      `Version mismatch: manifest=${manifestVersion}, package=${packageVersion}, lock=${lockVersion}, lock-root=${lockRootVersion}`
    );
  }

  const currentSemver = parseSemver(canonicalVersion, VERSION_FILE);
  const latestTag = readLatestTag();

  if (latestTag) {
    const latestSemver = parseSemver(latestTag, "latest git tag");
    const cmp = compareSemver(currentSemver, latestSemver);
    if (requireBump && cmp <= 0) {
      throw new Error(
        `Version must be greater than latest tag ${latestTag}. Current=${canonicalVersion}`
      );
    }
    if (cmp < 0) {
      throw new Error(
        `Version is older than latest tag ${latestTag}. Current=${canonicalVersion}`
      );
    }
  }

  console.log(
    `OK: VERSION=${canonicalVersion}${latestTag ? `, latest tag=${latestTag}` : ""}`
  );
} catch (error) {
  console.error(`VERSION_CHECK_FAILED: ${error.message}`);
  process.exit(1);
}
