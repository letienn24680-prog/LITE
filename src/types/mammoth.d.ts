declare module 'mammoth/mammoth.browser' {
  export interface MammothMessage {
    type: string;
    message: string;
  }

  export interface MammothResult {
    value: string;
    messages: MammothMessage[];
  }

  export interface MammothInput {
    arrayBuffer: ArrayBuffer;
  }

  export interface MammothImage {
    contentType: string;
    read: (encoding?: string) => Promise<string>;
  }

  export function convertToHtml(
    input: MammothInput,
    options?: Record<string, unknown>
  ): Promise<MammothResult>;

  export function extractRawText(input: MammothInput): Promise<MammothResult>;

  export const images: {
    imgElement: (
      fn: (image: MammothImage) => Promise<Record<string, string>>
    ) => unknown;
  };

  const mammoth: {
    convertToHtml: typeof convertToHtml;
    extractRawText: typeof extractRawText;
    images: typeof images;
  };

  export default mammoth;
}

declare module 'mammoth' {
  export interface MammothMessage {
    type: string;
    message: string;
  }

  export interface MammothResult {
    value: string;
    messages: MammothMessage[];
  }

  export interface MammothInput {
    arrayBuffer: ArrayBuffer;
  }

  export interface MammothImage {
    contentType: string;
    read: (encoding?: string) => Promise<string>;
  }

  export function convertToHtml(
    input: MammothInput,
    options?: Record<string, unknown>
  ): Promise<MammothResult>;

  export function extractRawText(input: MammothInput): Promise<MammothResult>;

  export const images: {
    imgElement: (
      fn: (image: MammothImage) => Promise<Record<string, string>>
    ) => unknown;
  };

  const mammoth: {
    convertToHtml: typeof convertToHtml;
    extractRawText: typeof extractRawText;
    images: typeof images;
  };

  export default mammoth;
}
