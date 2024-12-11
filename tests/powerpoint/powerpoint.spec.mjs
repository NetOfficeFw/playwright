import { powerpoint, expect } from '../../packages/playwright-mso-core/index.mjs'

powerpoint('test PowerPoint', async ({ page }) => {
  await page.screenshot({ path: 'screenshot.png' });
});
