/* global PpSlideLayout, MsoTextOrientation */
import { test, expect, delay } from '../../packages/playwright-mso-core';

test.describe('PowerPoint with Slido poll', () => {

  test.use({ filename: 'https://d.docs.live.net/8cd14a64b99957bc/Dokumenty/No%20AI%20Just%20Tests.pptx' });

  test('Slido', async ({ presentation, page }) => {
    await delay(2000);

    const title = await page.title();
    await expect(title).toEqual('Slido');

    await page.getByTestId('card-clickable-overlay').nth(5).click();
    await delay(1000);
    await page.getByTestId('pollQuestionTitle').fill('Word cloud poll');
    await page.getByTestId('submitPollFormButton').click();

    await delay(5000);
  })
});
