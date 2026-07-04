/** Plan and bucket tools. */

import { z } from "zod";
import type { ToolRegistrar } from "./registrar.js";
import type { GraphClient } from "../graph/client.js";
import type { PlannerBucket, PlannerPlan } from "../graph/types.js";
import { clip, failure, json, responseFormatField, text } from "./shared.js";

export function registerPlanTools(reg: ToolRegistrar, client: GraphClient): void {
  reg.register(
    {
      name: "planner_list_plans",
      title: "List Plans in a Group",
      description:
        "List all Planner plans owned by a Microsoft 365 group. Get the group ID from planner_search_groups.",
      inputSchema: {
        group_id: z.string().min(1).describe("Microsoft 365 group ID"),
        response_format: responseFormatField,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async (args: { group_id: string; response_format: "markdown" | "json" }) => {
      try {
        const plans = await client.getList<PlannerPlan>(`/groups/${args.group_id}/planner/plans`);
        if (plans.length === 0) return text("No plans found in this group.");
        if (args.response_format === "json") return text(clip(json(plans)));

        const lines = [`# Plans (${plans.length})`, ""];
        for (const p of plans) {
          lines.push(`- **${p.title}** (ID: ${p.id})${p.createdDateTime ? ` — created ${p.createdDateTime.slice(0, 10)}` : ""}`);
        }
        return text(clip(lines.join("\n")));
      } catch (error) {
        return failure(error);
      }
    }
  );

  reg.register(
    {
      name: "planner_get_plan",
      title: "Get a Plan with its Buckets",
      description: "Get a Planner plan and its buckets. Bucket IDs are needed to create or move tasks.",
      inputSchema: {
        plan_id: z.string().min(1).describe("Plan ID"),
        response_format: responseFormatField,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async (args: { plan_id: string; response_format: "markdown" | "json" }) => {
      try {
        const [plan, buckets] = await Promise.all([
          client.get<PlannerPlan>(`/planner/plans/${args.plan_id}`),
          client.getList<PlannerBucket>(`/planner/plans/${args.plan_id}/buckets`),
        ]);
        if (args.response_format === "json") return text(clip(json({ plan, buckets })));

        const lines = [`# ${plan.title}`, "", `- Plan ID: ${plan.id}`];
        if (plan.owner) lines.push(`- Owner group: ${plan.owner}`);
        lines.push("", `## Buckets (${buckets.length})`, "");
        for (const b of buckets) {
          lines.push(`- **${b.name}** (ID: ${b.id})`);
        }
        return text(clip(lines.join("\n")));
      } catch (error) {
        return failure(error);
      }
    }
  );

  reg.register(
    {
      name: "planner_create_bucket",
      title: "Create a Bucket",
      description: "Create a new bucket (column) in a plan.",
      inputSchema: {
        plan_id: z.string().min(1).describe("Plan ID"),
        name: z.string().min(1).describe("Bucket name"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
    },
    async (args: { plan_id: string; name: string }) => {
      try {
        const bucket = await client.post<PlannerBucket>("/planner/buckets", {
          planId: args.plan_id,
          name: args.name,
          orderHint: " !",
        });
        return text(`Bucket created: **${bucket.name}** (ID: ${bucket.id})`);
      } catch (error) {
        return failure(error);
      }
    }
  );
}
