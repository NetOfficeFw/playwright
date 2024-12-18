/* global PpSlideLayout, MsoTextOrientation */
import { test, expect, delay } from '../../packages/playwright-mso-core';

test('Prepare PowerPoint presentation using JavaScript code', async ({ presentation }) => {
  // create one slide in the presentation
  await presentation.evaluate(({ application, presentation }) => {
    const slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
    const textbox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 600, 100);
    textbox.TextFrame.TextRange.Text = `Slide ${slide.SlideID} in PowerPoint ${application.Build}`;
  });

  // await delay(1000);

  // await presentation.evaluate(({ application, presentation }) => {
  //   const slide = presentation.Slides.Add(2, PpSlideLayout.ppLayoutBlank);
  //   const textbox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 600, 100);
  //   textbox.TextFrame.TextRange.Text = `Hello World!`;
  //   slide.Select();
  // });

  await presentation.click('SlideShowFromCurrent');

  await delay(3000);
});