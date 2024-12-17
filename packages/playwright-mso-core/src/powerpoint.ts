import { test as base } from '@playwright/test';
import { ulid } from 'ulid';

import { PowerPointApp, Presentation } from './../types/index'
import { PowerPointAppProviderImpl } from './server/PowerPointAppProviderImpl';


type PowerPointTestFixture = {
  powerpoint: PowerPointApp;

  presentation: Presentation;
}

type PowerPointWorkerFixture = {
}

export const test = base.extend<PowerPointTestFixture, PowerPointWorkerFixture>({
  powerpoint: async ({}, use, testInfo) => {
    const sessionId = ulid();
    const port = 53080;
    var provider = new PowerPointAppProviderImpl(sessionId, port);

    // const powerpoint = await provider.launch();
    const powerpoint = await provider.connectOverGrpc(`http://127.0.0.1:${port}`);

    await use(powerpoint);
    // await powerpoint.close();
  },
  presentation: async ({ powerpoint }, use) => {
    const presentation = await powerpoint.newPresentation();
    if (presentation === null) {
      throw new Error('Failed to create new presentation.');
    }

    await use(presentation);
  },
});
