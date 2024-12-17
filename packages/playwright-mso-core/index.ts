export { test } from './src/powerpoint';
export { expect } from '@playwright/test';


export function delay(ms: number) {
  return new Promise(resolve => setTimeout(resolve, ms));
}
