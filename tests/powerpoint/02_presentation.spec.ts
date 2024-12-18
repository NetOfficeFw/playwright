import { test, expect, delay } from '../../packages/playwright-mso-core';

test('test presentation', async ({ presentation }) => {
  const title = await presentation.title();
  await expect(title).toEqual('Presentation2');

  await delay(2000);
});
