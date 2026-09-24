"use strict";

const fs = require("fs");

const SYSTEM_CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";

function launchOptions() {
  const executablePath = process.env.PLAYWRIGHT_EXECUTABLE_PATH || SYSTEM_CHROME;
  return fs.existsSync(executablePath)
    ? { headless: true, executablePath }
    : { headless: true };
}

module.exports = { launchOptions };
