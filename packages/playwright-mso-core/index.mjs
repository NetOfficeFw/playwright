import { test as base } from '@playwright/test';
import fs from 'fs';
import os from 'os';
import path from 'path';
import childProcess from 'child_process';

const EXECUTABLE_PATH = 'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE';

export const powerpoint = base.extend({
  browserVersion: ['131.0.0', { scope: 'worker' }],
  browserMajorVersion: [({ browserVersion }, use) => use(Number(browserVersion.split('.')[0])), { scope: 'worker' }],
  isAndroid: [false, { scope: 'worker' }],
  isElectron: [false, { scope: 'worker' }],
  electronMajorVersion: [0, { scope: 'worker' }],
  isWebView2: [true, { scope: 'worker' }],
  isHeadlessShell: [false, { scope: 'worker' }],

  browser: async ({ playwright }, use, testInfo) => {
    const cdpPort = 53080 + testInfo.workerIndex;
    // Make sure that the executable exists and is executable
    fs.accessSync(EXECUTABLE_PATH, fs.constants.X_OK);
    const userDataDir = path.join(
        fs.realpathSync.native(os.tmpdir()),
        `playwright-powerpoint-tests/user-data-dir-${testInfo.workerIndex}`,
    );
    const powerpointProcess = childProcess.spawn(EXECUTABLE_PATH, [], {
      shell: true,
      env: {
        ...process.env,
        NOPW_REMOTE_PORT: cdpPort,
        NOPW_USER_DATA_FOLDER: userDataDir,
      }
    });
    await new Promise(resolve => powerpointProcess.stdout.on('data', data => {
      if (data.toString().includes('NetOffice DevTools server running'))
        resolve();
    }));
    const browser = await playwright.chromium.connectOverCDP(`http://127.0.0.1:${cdpPort}`);
    await use(browser);
    await browser.close();
    childProcess.execSync(`taskkill /pid ${powerpointProcess.pid} /T /F`);
    fs.rmdirSync(userDataDir, { recursive: true });
  },
  context: async ({ browser }, use) => {
    const context = browser.contexts()[0];
    await use(context);
  },
  page: async ({ context }, use) => {
    const page = context.pages()[0];
    await use(page);
  },
});

export { expect } from '@playwright/test';
