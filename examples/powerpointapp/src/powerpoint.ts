import { test as base } from '@playwright/test';
import fs from 'fs';
import os from 'os';
import path from 'path';
import childProcess from 'child_process';

const EXECUTABLE_PATH = 'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE';

export const powerpoint = base.extend({
  browser: async ({ playwright }, use, testInfo) => {
    const cdpPort = 53080; // + testInfo.workerIndex;
    // Make sure that the executable exists and is executable
    fs.accessSync(EXECUTABLE_PATH, fs.constants.X_OK);
    const userDataDir = path.join(
        fs.realpathSync.native(os.tmpdir()),
        `playwright-powerpoint-tests/user-data-dir-${testInfo.workerIndex}`,
    );
    // const powerpointProcess = childProcess.spawn(`"${EXECUTABLE_PATH}"`, ['/B'], {
    //   shell: true,
    //   env: {
    //     ...process.env,
    //     PW_REMOTE_PORT: cdpPort,
    //     PW_USER_DATA_FOLDER: userDataDir,
    //   }
    // });
    // await new Promise(resolve => powerpointProcess.stdout.on('data', data => {
    //   const message = data.toString();
    //   console.log('PowerPoint', message);
    //   if (message.includes('NetOffice DevTools server running'))
    //     resolve();
    // }));
    const browser = await playwright.chromium.connectOverCDP(`http://127.0.0.1:${cdpPort}`);
    await use(browser);
    await browser.close();
    // childProcess.execSync(`taskkill /pid ${powerpointProcess.pid} /T /F`);

    try {
      fs.rmSync(userDataDir, { recursive: true, force: true });
    }
    catch (error) {
      if (error.code !== 'ENOENT') {
        throw error;
      }
    }
  },
  context: async ({ browser }, use) => {
    console.log('contexts', browser.contexts());
    const context = browser.contexts()[0];
    await use(context);
  },
  page: async ({ context }, use) => {
    console.log('pages', context.pages());
    const page = context.pages()[0];
    await use(page);
  },
});

export { expect } from '@playwright/test';
