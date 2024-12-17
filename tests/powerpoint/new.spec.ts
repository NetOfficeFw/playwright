import { test, expect, delay } from '../../packages/playwright-mso-core';

test('test Microsoft PowerPoint app', async ({ powerpoint }) => {
  await expect(powerpoint.version()).toEqual('16.0.18330');
});

test('test presentation', async ({ presentation }) => {
  const title = await presentation.title();
  await expect(title).toEqual('Presentation1.pptx');
});


test('Prepare PowerPoint presentation using JavaScript code', async ({ presentation }) => {
  // create one slide in the presentation
  await presentation.evaluate(({ application, presentation }) => {
    const slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
    const textbox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 600, 100);
    textbox.TextFrame.TextRange.Text = `Slide ${slide.SlideID} in PowerPoint ${application.Build}`;
  });

  await delay(1000);

  await presentation.evaluate(({ application, presentation }) => {
    const slide = presentation.Slides.Add(2, PpSlideLayout.ppLayoutBlank);
    const textbox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 600, 100);
    textbox.TextFrame.TextRange.Text = `Hello World!`;
    slide.Select();
  });

  await delay(3000);
});
