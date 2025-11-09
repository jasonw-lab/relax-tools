declare module 'encoding-japanese' {
  interface ConvertOptions {
    to: string;
    from: string;
    type?: 'string' | 'array';
  }

  interface Encoding {
    Convert(data: Uint8Array | number[], options: ConvertOptions): string | number[];
    codeToString(code: number[] | Uint8Array): string;
    stringToCode(str: string): number[];
  }

  const Encoding: Encoding;
  export default Encoding;
}

