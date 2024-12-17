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

export type PresentationEvaluateArgs = {
  application?: any,
  presentation?: any
}

export type PageFunction = ((args: PresentationEvaluateArgs) => void | Promise<void>);

export interface Presentation {
  title(): Promise<string>;
  fullname(): Promise<string>;

  evaluate(script: PageFunction): Promise<void>;
}


export enum msoTextOrientation {
  Horizontal = 1,
  Vertical = 5
}
