import { powerpoint, expect } from './../src/powerpoint';

powerpoint('test Microsoft PowerPoint', async ({ page }) => {
  // await page.screenshot({ path: 'screenshot.png' });

  await page.goto('https://playwright.dev');
  const getStarted = page.getByText('Get Started');
  await expect(getStarted).toBeVisible();
});
