import * as fs from "fs";
import * as path from "path";

let logPath: string | null = null;

export function initLogger(filePath: string): void {
  logPath = filePath;
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
}

function write(level: "INFO" | "WARN" | "ERROR", message: string): void {
  const line = `[${new Date().toISOString()}] [${level}] ${message}`;
  if (level === "ERROR") {
    console.error(line);
  } else {
    console.log(line);
  }
  if (logPath) {
    fs.appendFileSync(logPath, line + "\n");
  }
}

export const logger = {
  info: (message: string) => write("INFO", message),
  warn: (message: string) => write("WARN", message),
  error: (message: string) => write("ERROR", message),
};
