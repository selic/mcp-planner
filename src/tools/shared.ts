/** Shared zod fragments and result helpers for tool implementations. */

import { z } from "zod";
import { describeError } from "../graph/client.js";

/** Hard cap on tool response size, to protect the model's context window. */
export const RESPONSE_CHAR_LIMIT = 25_000;

export const responseFormatField = z
  .enum(["markdown", "json"])
  .default("markdown")
  .describe("Output format: human-readable markdown (default) or structured JSON");

export const topField = z
  .number()
  .int()
  .positive()
  .max(200)
  .default(50)
  .describe("Maximum results to return (default 50, max 200)");

export interface ToolResult {
  content: Array<{ type: "text"; text: string }>;
  isError?: boolean;
}

export function text(value: string): ToolResult {
  return { content: [{ type: "text", text: value }] };
}

export function failure(error: unknown): ToolResult {
  return { content: [{ type: "text", text: describeError(error) }], isError: true };
}

export function json(value: unknown): string {
  return JSON.stringify(value, null, 2);
}

export function clip(value: string, hint?: string): string {
  if (value.length <= RESPONSE_CHAR_LIMIT) return value;
  const note = hint ?? "Use filters or a smaller top value to narrow the result.";
  return (
    value.slice(0, RESPONSE_CHAR_LIMIT) +
    `\n\n---\n[Truncated at ${RESPONSE_CHAR_LIMIT.toLocaleString()} characters. ${note}]`
  );
}
