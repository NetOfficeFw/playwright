import fs from 'fs';
import os from 'os';
import path from 'path';
import childProcess from 'child_process';

import type * as api from "./../../types";

const EXECUTABLE_PATH = 'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE';

export class PowerPointAppProviderImpl implements api.PowerPointAppProvider {
  #port: number;
  sessionId: string;

  constructor(
    sessionId: string,
    port: number
  ) {
    this.sessionId = sessionId;
    this.#port = port;
  }

  async connectOverGrpc(endpointURL: string): Promise<api.PowerPointApp> {
    const app = new PowerPointApp();
    await app.connectOverGrpc(endpointURL);
    return app;
  }

  async launch(filename: string | null): Promise<api.PowerPointApp> {
    const expectedRunningMessage = `NetOffice DevTools server running session ${this.sessionId}`;

    const args = filename ? ['/O', filename] : ['/B'];

    // Make sure that the executable exists and is executable
    fs.accessSync(EXECUTABLE_PATH, fs.constants.X_OK);
    const powerpointProcess = childProcess.spawn(`"${EXECUTABLE_PATH}"`, args, {
      shell: true,
      env: {
        ...process.env,
        PW_SESSION_ID: this.sessionId,
        PW_GRPC_PORT: String(this.#port)
      }
    });
    await new Promise<void>(resolve => powerpointProcess.stdout.on('data', data => {
      const message = data.toString();
      if (message.includes(expectedRunningMessage))
        setTimeout(() => resolve(), 500);
    }));

    const app = await this.connectOverGrpc(`http://127.0.0.1:${this.#port}`);
    return app;
  }
}

class PowerPointApp implements api.PowerPointApp {
  #appType: string;
  #version: string;
  #pid: number;
  #endpoint: string;

  constructor() {
    this.#appType = 'powerpoint';
    this.#version = '0.0';
    this.#pid = 0;
    this.#endpoint = '';
  }

  async connectOverGrpc(endpointURL: string): Promise<void> {
    this.#endpoint = endpointURL;

    const getVersion = `${endpointURL}/json/version`;
    const response = await fetch(getVersion);
    if (response.ok) {
      const data = await response.json();
      // console.log(data);

      this.#appType = data.app_type;
      this.#version = data.version;
      this.#pid = data.process_id;
    }
  }

  appType(): string {
    return this.#appType;
  }

  version(): string {
    return this.#version;
  }

  async newPresentation(filename: string | null): Promise<api.Presentation | null> {
    const query = filename ? `?filename=${encodeURIComponent(filename)}` : '';
    const newPresentationEndpoint = `${this.#endpoint}/newPresentation${query}`;
    console.log(newPresentationEndpoint);
    const response = await fetch(newPresentationEndpoint, { method: 'POST' });
    if (response.ok) {
      var data = await response.json();
      // console.log(data);
      return new Presentation(this.#endpoint, data.title, data.fullname);
    }

    return null;
  }

  async close(): Promise<void> {
    const closeEndpoint = `${this.#endpoint}/close`;
    const response = await fetch(closeEndpoint, { method: 'POST' });
    if (response.ok) {
      // console.log(await response.json());
    }
  }

  async [Symbol.asyncDispose](): Promise<void> {
    console.log('Disposing PowerPoint channel.');
  }
}

class Presentation implements api.Presentation {
  #endpoint: string;
  #title: string;
  #fullname: string;

  constructor(endpoint: string, title: string, fullname: string) {
    this.#endpoint = endpoint;
    this.#title = title;
    this.#fullname = fullname;
  }

  async title(): Promise<string> {
    return this.#title;
  }

  async fullname(): Promise<string> {
    return this.#fullname;
  }

  async evaluate(script: api.PageFunction): Promise<void> {
    const scriptText = script.toString();

    const evaluateEndpoint = `${this.#endpoint}/evaluate`;
    const response = await fetch(evaluateEndpoint, { method: 'POST', body: scriptText, headers: { 'Content-Type': 'text/javascript' } });
    if (response.ok) {
      // console.log('Script was evaluated.');
    }
  }

  async click(msoId: string): Promise<void> {
    const evaluateEndpoint = `${this.#endpoint}/click/${msoId}`;
    const response = await fetch(evaluateEndpoint, { method: 'POST' });
  }
}
