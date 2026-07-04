/** Task tools — list, read, create, update, complete, delete. */

import { z } from "zod";
import type { ToolRegistrar } from "./registrar.js";
import type { GraphClient } from "../graph/client.js";
import {
  PRIORITY_VALUES,
  priorityLabel,
  progressLabel,
  type PlannerAssignments,
  type PlannerTask,
  type PlannerTaskDetails,
  type PriorityName,
} from "../graph/types.js";
import { clip, failure, json, responseFormatField, text, topField, type ToolResult } from "./shared.js";

const ASSIGNMENT_TYPE = "#microsoft.graph.plannerAssignment";
const CHECKLIST_TYPE = "#microsoft.graph.plannerChecklistItem";

function taskLine(t: PlannerTask): string {
  const parts = [`- **${t.title}** (ID: ${t.id})`];
  parts.push(`  ${progressLabel(t.percentComplete)}, priority ${priorityLabel(t.priority)}`);
  if (t.dueDateTime) parts.push(`, due ${t.dueDateTime.slice(0, 10)}`);
  const assignees = Object.keys(t.assignments ?? {});
  if (assignees.length > 0) parts.push(`, assignees: ${assignees.join(", ")}`);
  return parts.join("");
}

function renderTaskList(title: string, tasks: PlannerTask[], format: "markdown" | "json"): ToolResult {
  if (tasks.length === 0) return text("No tasks found.");
  if (format === "json") return text(clip(json(tasks)));
  const lines = [`# ${title} (${tasks.length})`, ""];
  for (const t of tasks) lines.push(taskLine(t));
  return text(clip(lines.join("\n")));
}

export function registerTaskTools(reg: ToolRegistrar, client: GraphClient): void {
  reg.register(
    {
      name: "planner_list_tasks",
      title: "List Tasks in a Plan or Bucket",
      description:
        "List Planner tasks in a plan (or a single bucket). Optionally filter by assignee user ID or open/completed state.",
      inputSchema: {
        plan_id: z.string().optional().describe("Plan ID (required unless bucket_id is given)"),
        bucket_id: z.string().optional().describe("Bucket ID — list only this bucket's tasks"),
        assigned_to: z.string().optional().describe("Only tasks assigned to this user ID"),
        include_completed: z.boolean().default(false).describe("Include completed tasks (default false)"),
        top: topField,
        response_format: responseFormatField,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async (args: {
      plan_id?: string;
      bucket_id?: string;
      assigned_to?: string;
      include_completed: boolean;
      top: number;
      response_format: "markdown" | "json";
    }) => {
      try {
        if (!args.plan_id && !args.bucket_id) {
          return text("Error: provide plan_id or bucket_id.");
        }
        const path = args.bucket_id
          ? `/planner/buckets/${args.bucket_id}/tasks`
          : `/planner/plans/${args.plan_id}/tasks`;
        let tasks = await client.getList<PlannerTask>(path);
        if (!args.include_completed) tasks = tasks.filter((t) => t.percentComplete !== 100);
        if (args.assigned_to) tasks = tasks.filter((t) => Object.keys(t.assignments ?? {}).includes(args.assigned_to!));
        tasks = tasks.slice(0, args.top);
        return renderTaskList("Tasks", tasks, args.response_format);
      } catch (error) {
        return failure(error);
      }
    }
  );

  reg.register(
    {
      name: "planner_list_user_tasks",
      title: "List a User's Tasks",
      description:
        "List all Planner tasks assigned to a user across every plan they participate in. Get the user ID from planner_find_user.",
      inputSchema: {
        user_id: z.string().min(1).describe("User ID"),
        include_completed: z.boolean().default(false).describe("Include completed tasks (default false)"),
        top: topField,
        response_format: responseFormatField,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async (args: { user_id: string; include_completed: boolean; top: number; response_format: "markdown" | "json" }) => {
      try {
        let tasks = await client.getList<PlannerTask>(`/users/${args.user_id}/planner/tasks`);
        if (!args.include_completed) tasks = tasks.filter((t) => t.percentComplete !== 100);
        tasks = tasks.slice(0, args.top);
        return renderTaskList("User tasks", tasks, args.response_format);
      } catch (error) {
        return failure(error);
      }
    }
  );

  reg.register(
    {
      name: "planner_get_task",
      title: "Get a Task with Details",
      description:
        "Get a Planner task including its description and checklist. Also returns the etags needed by update tools (updates fetch them automatically).",
      inputSchema: {
        task_id: z.string().min(1).describe("Task ID"),
        response_format: responseFormatField,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async (args: { task_id: string; response_format: "markdown" | "json" }) => {
      try {
        const [task, details] = await Promise.all([
          client.get<PlannerTask>(`/planner/tasks/${args.task_id}`),
          client.get<PlannerTaskDetails>(`/planner/tasks/${args.task_id}/details`),
        ]);
        if (args.response_format === "json") return text(clip(json({ task, details })));

        const lines = [`# ${task.title}`, ""];
        lines.push(`- ID: ${task.id}`);
        lines.push(`- Plan: ${task.planId}${task.bucketId ? ` | Bucket: ${task.bucketId}` : ""}`);
        lines.push(`- Status: ${progressLabel(task.percentComplete)} | Priority: ${priorityLabel(task.priority)}`);
        if (task.startDateTime) lines.push(`- Start: ${task.startDateTime.slice(0, 10)}`);
        if (task.dueDateTime) lines.push(`- Due: ${task.dueDateTime.slice(0, 10)}`);
        if (task.completedDateTime) lines.push(`- Completed: ${task.completedDateTime.slice(0, 10)}`);
        const assignees = Object.keys(task.assignments ?? {});
        if (assignees.length > 0) lines.push(`- Assignee user IDs: ${assignees.join(", ")}`);
        if (details.description) lines.push("", "## Description", "", details.description);
        const checklist = Object.entries(details.checklist ?? {});
        if (checklist.length > 0) {
          lines.push("", "## Checklist", "");
          for (const [id, item] of checklist) {
            if (item) lines.push(`- [${item.isChecked ? "x" : " "}] ${item.title} (item ID: ${id})`);
          }
        }
        return text(clip(lines.join("\n")));
      } catch (error) {
        return failure(error);
      }
    }
  );

  reg.register(
    {
      name: "planner_create_task",
      title: "Create a Task",
      description:
        "Create a Planner task in a plan. Optionally set the bucket, due date, priority, assignees, and a description.",
      inputSchema: {
        plan_id: z.string().min(1).describe("Plan ID"),
        title: z.string().min(1).describe("Task title"),
        bucket_id: z.string().optional().describe("Bucket ID (defaults to the plan's default bucket)"),
        due_date: z.string().optional().describe("Due date, ISO format e.g. 2026-07-31"),
        priority: z.enum(["urgent", "important", "medium", "low"]).optional().describe("Task priority"),
        assignee_ids: z.array(z.string()).optional().describe("User IDs to assign (from planner_find_user)"),
        description: z.string().optional().describe("Task description (notes)"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
    },
    async (args: {
      plan_id: string;
      title: string;
      bucket_id?: string;
      due_date?: string;
      priority?: PriorityName;
      assignee_ids?: string[];
      description?: string;
    }) => {
      try {
        const assignments: PlannerAssignments = {};
        for (const id of args.assignee_ids ?? []) {
          assignments[id] = { "@odata.type": ASSIGNMENT_TYPE, orderHint: " !" };
        }
        const task = await client.post<PlannerTask>("/planner/tasks", {
          planId: args.plan_id,
          title: args.title,
          ...(args.bucket_id ? { bucketId: args.bucket_id } : {}),
          ...(args.due_date ? { dueDateTime: `${args.due_date}T17:00:00Z` } : {}),
          ...(args.priority ? { priority: PRIORITY_VALUES[args.priority] } : {}),
          ...(Object.keys(assignments).length > 0 ? { assignments } : {}),
        });

        if (args.description) {
          // Details are a separate resource with their own etag; fetch then patch.
          const details = await client.get<PlannerTaskDetails>(`/planner/tasks/${task.id}/details`);
          await client.patch(`/planner/tasks/${task.id}/details`, { description: args.description }, details["@odata.etag"]!);
        }
        return text(`Task created: **${task.title}** (ID: ${task.id})`);
      } catch (error) {
        return failure(error);
      }
    }
  );

  reg.register(
    {
      name: "planner_update_task",
      title: "Update a Task",
      description:
        "Update a Planner task: title, bucket (move), due date, priority, progress (not started / in progress / completed), or assignees. Fetches the current etag automatically.",
      inputSchema: {
        task_id: z.string().min(1).describe("Task ID"),
        title: z.string().optional().describe("New title"),
        bucket_id: z.string().optional().describe("Move to this bucket ID"),
        due_date: z.string().optional().describe("New due date (ISO, e.g. 2026-07-31) or empty string to clear"),
        priority: z.enum(["urgent", "important", "medium", "low"]).optional().describe("New priority"),
        progress: z.enum(["not_started", "in_progress", "completed"]).optional().describe("New progress state"),
        add_assignee_ids: z.array(z.string()).optional().describe("User IDs to assign"),
        remove_assignee_ids: z.array(z.string()).optional().describe("User IDs to unassign"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async (args: {
      task_id: string;
      title?: string;
      bucket_id?: string;
      due_date?: string;
      priority?: PriorityName;
      progress?: "not_started" | "in_progress" | "completed";
      add_assignee_ids?: string[];
      remove_assignee_ids?: string[];
    }) => {
      try {
        const current = await client.get<PlannerTask>(`/planner/tasks/${args.task_id}`);
        const body: Record<string, unknown> = {};
        if (args.title !== undefined) body.title = args.title;
        if (args.bucket_id !== undefined) body.bucketId = args.bucket_id;
        if (args.due_date !== undefined) {
          body.dueDateTime = args.due_date === "" ? null : `${args.due_date}T17:00:00Z`;
        }
        if (args.priority !== undefined) body.priority = PRIORITY_VALUES[args.priority];
        if (args.progress !== undefined) {
          body.percentComplete = args.progress === "completed" ? 100 : args.progress === "in_progress" ? 50 : 0;
        }
        const assignments: PlannerAssignments = {};
        for (const id of args.add_assignee_ids ?? []) {
          assignments[id] = { "@odata.type": ASSIGNMENT_TYPE, orderHint: " !" };
        }
        for (const id of args.remove_assignee_ids ?? []) assignments[id] = null;
        if (Object.keys(assignments).length > 0) body.assignments = assignments;

        if (Object.keys(body).length === 0) return text("Nothing to update — provide at least one field.");

        const updated = await client.patch<PlannerTask>(
          `/planner/tasks/${args.task_id}`,
          body,
          current["@odata.etag"]!
        );
        return text(
          `Task updated: **${updated.title ?? current.title}** — ${progressLabel(updated.percentComplete)}, priority ${priorityLabel(updated.priority)}`
        );
      } catch (error) {
        return failure(error);
      }
    }
  );

  reg.register(
    {
      name: "planner_update_task_details",
      title: "Update Task Description or Checklist",
      description:
        "Update a task's description, add checklist items, or check/uncheck existing checklist items (item IDs come from planner_get_task).",
      inputSchema: {
        task_id: z.string().min(1).describe("Task ID"),
        description: z.string().optional().describe("New description (replaces the existing one)"),
        add_checklist_items: z.array(z.string()).optional().describe("Checklist item titles to add"),
        check_item_ids: z.array(z.string()).optional().describe("Checklist item IDs to mark checked"),
        uncheck_item_ids: z.array(z.string()).optional().describe("Checklist item IDs to mark unchecked"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async (args: {
      task_id: string;
      description?: string;
      add_checklist_items?: string[];
      check_item_ids?: string[];
      uncheck_item_ids?: string[];
    }) => {
      try {
        const details = await client.get<PlannerTaskDetails>(`/planner/tasks/${args.task_id}/details`);
        const body: Record<string, unknown> = {};
        if (args.description !== undefined) body.description = args.description;

        const checklist: Record<string, unknown> = {};
        for (const title of args.add_checklist_items ?? []) {
          checklist[crypto.randomUUID()] = { "@odata.type": CHECKLIST_TYPE, title, isChecked: false };
        }
        for (const id of args.check_item_ids ?? []) {
          checklist[id] = { "@odata.type": CHECKLIST_TYPE, isChecked: true };
        }
        for (const id of args.uncheck_item_ids ?? []) {
          checklist[id] = { "@odata.type": CHECKLIST_TYPE, isChecked: false };
        }
        if (Object.keys(checklist).length > 0) body.checklist = checklist;

        if (Object.keys(body).length === 0) return text("Nothing to update — provide at least one field.");

        await client.patch(`/planner/tasks/${args.task_id}/details`, body, details["@odata.etag"]!);
        return text("Task details updated.");
      } catch (error) {
        return failure(error);
      }
    }
  );

  reg.register(
    {
      name: "planner_delete_task",
      title: "Delete a Task",
      description: "Permanently delete a Planner task. This cannot be undone.",
      inputSchema: {
        task_id: z.string().min(1).describe("Task ID"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: true },
    },
    async (args: { task_id: string }) => {
      try {
        const current = await client.get<PlannerTask>(`/planner/tasks/${args.task_id}`);
        await client.delete(`/planner/tasks/${args.task_id}`, current["@odata.etag"]!);
        return text(`Task deleted: ${current.title} (${args.task_id})`);
      } catch (error) {
        return failure(error);
      }
    }
  );
}
