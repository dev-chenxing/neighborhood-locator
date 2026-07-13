export type 街道号数映射 = {
  [street: string]: {
    name: string;
    start?: number;
    end?: number;
    oddity?: number;
    all?: boolean;
  }[];
};

export type LogLevel = "INFO" | "ERROR";
