import { powerpoint, expect } from './../../packages/playwright-mso-core/index.mjs';

powerpoint('test Microsoft PowerPoint', async ({ page }) => {
  // await page.screenshot({ path: 'screenshot.png' });

  await page.goto('https://playwright.dev');
  const getStarted = page.getByText('Get Started');
  await expect(getStarted).toBeVisible();
});
