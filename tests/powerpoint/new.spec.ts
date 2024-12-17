import { test, expect } from '../../packages/playwright-mso-core';

test('test Microsoft PowerPoint app', async ({ powerpoint }) => {
  await expect(powerpoint.version()).toEqual('16.0.18330');
});

test('test presentation', async ({ presentation }) => {
  const title = await presentation.title();
  await expect(title).toEqual('Presentation1.pptx');
});
