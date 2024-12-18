import { test, expect, delay } from '../../packages/playwright-mso-core';

test('test Microsoft PowerPoint app', async ({ powerpoint }) => {
  await expect(powerpoint.version()).toEqual('16.0.18330');

  await delay(2000);
});
