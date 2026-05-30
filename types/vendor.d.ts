declare module "papaparse" {
  export type ParseError = {
    code: string;
    message: string;
    row?: number;
    type: string;
  };

  export type ParseResult<T> = {
    data: T[];
    errors: ParseError[];
    meta: unknown;
  };

  export type ParseConfig<T> = {
    header?: boolean;
    skipEmptyLines?: boolean | "greedy";
    dynamicTyping?: boolean;
    complete?: (results: ParseResult<T>) => void;
  };

  const Papa: {
    parse<T>(input: string, config?: ParseConfig<T>): ParseResult<T>;
  };

  export default Papa;
}
