/**
 * Microsoft Graph client for Planner on the global fetch API.
 *
 * - Base URL: https://graph.microsoft.com/v1.0
 * - Auth: OAuth2 client-credentials flow against
 *   https://login.microsoftonline.com/<tenant>/oauth2/v2.0/token with the
 *   .default scope. Tokens are cached until shortly before expiry.
 * - Planner concurrency: every PATCH/DELETE on planner resources requires an
 *   If-Match header carrying the resource's @odata.etag. `patch`/`delete`
 *   take the etag explicitly; tools fetch the current resource first.
 * - PATCH sends `Prefer: return=representation` so updates return the new state.
 */

const REQUEST_TIMEOUT_MS = 30_000;
const TOKEN_SKEW_MS = 60_000;

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

export class GraphApiError extends Error {
  constructor(
    message: string,
    public readonly status?: number,
    public readonly detail?: string
  ) {
    super(message);
    this.name = "GraphApiError";
  }
}

export function describeError(error: unknown): string {
  if (error instanceof GraphApiError) {
    const detail = error.detail ? ` ${error.detail}` : "";
    switch (error.status) {
      case 400:
        return `Error: Bad request.${detail} Check the parameters.`;
      case 401:
        return "Error: Microsoft Graph authentication failed — check the tenant id, client id and client secret.";
      case 403:
        return `Error: Permission denied.${detail} The app registration may be missing an application permission (Tasks.ReadWrite.All, GroupMember.Read.All, User.Read.All) or admin consent.`;
      case 404:
        return "Error: Resource not found. Verify the ID is correct.";
      case 409:
        return `Error: Conflict.${detail}`;
      case 412:
        return "Error: The resource changed since it was read (ETag mismatch). Fetch it again and retry.";
      case 429:
        return "Error: Microsoft Graph rate limit hit. Wait a moment and retry.";
      default:
        if (error.status && error.status >= 500) {
          return `Error: Microsoft Graph server error (${error.status}). Try again later.`;
        }
        return `Error: Graph request failed${error.status ? ` with status ${error.status}` : ""}.${detail}`;
    }
  }
  if (error instanceof Error && error.name === "TimeoutError") {
    return "Error: Microsoft Graph request timed out. Try again.";
  }
  return `Error: ${error instanceof Error ? error.message : String(error)}`;
}

export interface GraphCredentials {
  tenantId: string;
  clientId: string;
  clientSecret: string;
}

interface TokenResponse {
  access_token: string;
  expires_in: number;
}

export class GraphClient {
  private token: string | undefined;
  private tokenExpiresAt = 0;

  constructor(private readonly credentials: GraphCredentials) {}

  private async getToken(): Promise<string> {
    if (this.token && Date.now() < this.tokenExpiresAt - TOKEN_SKEW_MS) return this.token;

    const url = `https://login.microsoftonline.com/${this.credentials.tenantId}/oauth2/v2.0/token`;
    let response: Response;
    try {
      response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type: "client_credentials",
          client_id: this.credentials.clientId,
          client_secret: this.credentials.clientSecret,
          scope: "https://graph.microsoft.com/.default",
        }),
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
    } catch (err) {
      if (err instanceof Error && err.name === "TimeoutError") throw err;
      throw new GraphApiError(
        `Could not reach the Microsoft login endpoint: ${err instanceof Error ? err.message : String(err)}`
      );
    }

    if (!response.ok) {
      let detail: string | undefined;
      try {
        const body = (await response.json()) as { error_description?: string; error?: string };
        detail = body.error_description ?? body.error;
      } catch {
        // non-JSON error body
      }
      throw new GraphApiError("Token request failed", 401, detail);
    }

    const body = (await response.json()) as TokenResponse;
    this.token = body.access_token;
    this.tokenExpiresAt = Date.now() + body.expires_in * 1000;
    return this.token;
  }

  private async request<T>(
    method: "GET" | "POST" | "PATCH" | "DELETE",
    path: string,
    options: {
      query?: Record<string, string | number | undefined>;
      body?: unknown;
      etag?: string;
      /** Extra headers, e.g. ConsistencyLevel: eventual for $search/$count. */
      headers?: Record<string, string>;
    } = {}
  ): Promise<T> {
    const token = await this.getToken();
    const url = new URL(GRAPH_BASE + path);
    for (const [key, value] of Object.entries(options.query ?? {})) {
      if (value !== undefined) url.searchParams.set(key, String(value));
    }

    let response: Response;
    try {
      response = await fetch(url, {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
          ...(options.body !== undefined ? { "Content-Type": "application/json" } : {}),
          ...(options.etag ? { "If-Match": options.etag } : {}),
          ...(method === "PATCH" ? { Prefer: "return=representation" } : {}),
          ...options.headers,
        },
        body: options.body === undefined ? undefined : JSON.stringify(options.body),
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
    } catch (err) {
      if (err instanceof Error && err.name === "TimeoutError") throw err;
      throw new GraphApiError(
        `Could not reach Microsoft Graph: ${err instanceof Error ? err.message : String(err)}`
      );
    }

    if (!response.ok) {
      let detail: string | undefined;
      try {
        const errBody = (await response.json()) as { error?: { message?: string; code?: string } };
        detail = errBody.error?.message ?? errBody.error?.code;
      } catch {
        // non-JSON error body
      }
      throw new GraphApiError(`Graph responded with ${response.status}`, response.status, detail);
    }

    if (response.status === 204) return undefined as T;
    return (await response.json()) as T;
  }

  async get<T>(path: string, query?: Record<string, string | number | undefined>, headers?: Record<string, string>): Promise<T> {
    return this.request<T>("GET", path, { query, headers });
  }

  /** GET a collection endpoint, returning `value` and following no pages (Planner collections are small). */
  async getList<T>(path: string, query?: Record<string, string | number | undefined>, headers?: Record<string, string>): Promise<T[]> {
    const result = await this.request<{ value: T[] }>("GET", path, { query, headers });
    return result.value;
  }

  async post<T>(path: string, body: unknown): Promise<T> {
    return this.request<T>("POST", path, { body });
  }

  async patch<T>(path: string, body: unknown, etag: string): Promise<T> {
    return this.request<T>("PATCH", path, { body, etag });
  }

  async delete(path: string, etag: string): Promise<void> {
    await this.request<void>("DELETE", path, { etag });
  }
}
