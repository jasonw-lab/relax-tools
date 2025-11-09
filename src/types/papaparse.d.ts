declare module 'papaparse' {
  interface ParseConfig<T = any> {
    header?: boolean;
    skipEmptyLines?: boolean;
    delimiter?: string;
    newline?: string;
    quoteChar?: string;
    escapeChar?: string;
    comments?: boolean | string;
    complete?: (results: ParseResult<T>) => void;
    error?: (error: ParseError) => void;
  }

  interface ParseResult<T = any> {
    data: T[];
    errors: ParseError[];
    meta: {
      delimiter: string;
      linebreak: string;
      aborted: boolean;
      fields?: string[];
      truncated: boolean;
    };
  }

  interface ParseError {
    type: string;
    code: string;
    message: string;
    row?: number;
  }

  interface PapaParse {
    parse<T = any>(input: string | File, config?: ParseConfig<T>): ParseResult<T>;
  }

  const Papa: PapaParse;
  export default Papa;
}

