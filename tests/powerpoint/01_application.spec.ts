import { test, expect, delay } from '../../packages/playwright-mso-core';

test('test Microsoft PowerPoint app', async ({ powerpoint }) => {
  await expect(powerpoint.version()).toEqual('16.0.18330');

  await delay(2000);
});

test('browser test', async ({ page }) => {
  const title = await page.title();
  await expect(title).toEqual('Slido');

//   await page.getByTestId('loginButton').click();

  await page.getByTestId('card-clickable-overlay').nth(4).click();
  await page.getByTestId('pollQuestionTitle').fill('Open text poll');
  await page.getByTestId('submitPollFormButton').click();
});
