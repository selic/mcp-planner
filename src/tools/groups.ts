/** Group and user discovery tools — find the IDs the Planner tools need. */

import { z } from "zod";
import type { ToolRegistrar } from "./registrar.js";
import type { GraphClient } from "../graph/client.js";
import type { Group, User } from "../graph/types.js";
import { clip, failure, json, responseFormatField, text, topField } from "./shared.js";

export function registerGroupTools(reg: ToolRegistrar, client: GraphClient): void {
  reg.register(
    {
      name: "planner_search_groups",
      title: "Search Microsoft 365 Groups",
      description:
        "Search Microsoft 365 groups (Teams) by display name. Use this first to find the group ID required by planner_list_plans.",
      inputSchema: {
        name_contains: z.string().min(1).describe("Text the group display name contains"),
        top: topField,
        response_format: responseFormatField,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async (args: { name_contains: string; top: number; response_format: "markdown" | "json" }) => {
      try {
        const groups = await client.getList<Group>(
          "/groups",
          {
            $search: `"displayName:${args.name_contains.replace(/"/g, "")}"`,
            $select: "id,displayName,mail,description",
            $top: args.top,
          },
          { ConsistencyLevel: "eventual" }
        );
        if (groups.length === 0) return text("No groups found.");
        if (args.response_format === "json") return text(clip(json(groups)));

        const lines = [`# Groups (${groups.length})`, ""];
        for (const g of groups) {
          lines.push(`- **${g.displayName}** (ID: ${g.id})${g.mail ? ` — ${g.mail}` : ""}`);
        }
        return text(clip(lines.join("\n")));
      } catch (error) {
        return failure(error);
      }
    }
  );

  reg.register(
    {
      name: "planner_find_user",
      title: "Find a User",
      description:
        "Find a user by display name or UPN/email to get the user ID needed for task assignments and planner_list_user_tasks.",
      inputSchema: {
        query: z.string().min(1).describe("Display name fragment or userPrincipalName/email"),
        top: topField,
        response_format: responseFormatField,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async (args: { query: string; top: number; response_format: "markdown" | "json" }) => {
      try {
        const escaped = args.query.replace(/'/g, "''");
        const users = await client.getList<User>(
          "/users",
          {
            $filter: `startswith(displayName,'${escaped}') or startswith(userPrincipalName,'${escaped}')`,
            $select: "id,displayName,userPrincipalName",
            $top: args.top,
          }
        );
        if (users.length === 0) return text("No users found.");
        if (args.response_format === "json") return text(clip(json(users)));

        const lines = [`# Users (${users.length})`, ""];
        for (const u of users) {
          lines.push(`- **${u.displayName}** (ID: ${u.id}) — ${u.userPrincipalName}`);
        }
        return text(clip(lines.join("\n")));
      } catch (error) {
        return failure(error);
      }
    }
  );
}
