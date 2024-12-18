/* global PpSlideLayout, MsoTextOrientation */
import { test, expect, delay } from '../../packages/playwright-mso-core';

test('Create poll in a Sidebar', async ({ page }) => {
  const title = await page.title();
  await expect(title).toEqual('Slido');

  await page.getByTestId('card-clickable-overlay').nth(5).click();
  await page.getByTestId('pollQuestionTitle').fill('Word cloud poll');
  await page.getByTestId('submitPollFormButton').click();
});
