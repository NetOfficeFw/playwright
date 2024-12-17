export interface PowerPointAppProvider {
  sessionId: string;

  launch(): Promise<PowerPointApp>;

  connectOverGrpc(endpointURL: string): Promise<PowerPointApp>;
}


export interface PowerPointApp {
  appType(): string;

  version(): string;

  newPresentation(): Promise<Presentation | null>;

  close(): Promise<void>;
  [Symbol.asyncDispose](): Promise<void>;
}

export interface Presentation {
  title(): Promise<string>;
  fullname(): Promise<string>;
}
