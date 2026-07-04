/** Microsoft Graph resource shapes (fields this server actually uses). */

export interface Group {
  id: string;
  displayName: string;
  mail?: string;
  description?: string;
}

export interface User {
  id: string;
  displayName: string;
  userPrincipalName: string;
}

export interface PlannerPlan {
  "@odata.etag"?: string;
  id: string;
  title: string;
  owner?: string; // group id
  createdDateTime?: string;
}

export interface PlannerBucket {
  "@odata.etag"?: string;
  id: string;
  name: string;
  planId: string;
  orderHint?: string;
}

export interface PlannerAssignments {
  [userId: string]: {
    "@odata.type"?: string;
    orderHint?: string;
    assignedDateTime?: string;
  } | null;
}

export interface PlannerTask {
  "@odata.etag"?: string;
  id: string;
  planId: string;
  bucketId?: string;
  title: string;
  percentComplete?: number; // 0 | 50 | 100
  priority?: number; // 0-10: 1 urgent, 3 important, 5 medium, 9 low
  dueDateTime?: string | null;
  startDateTime?: string | null;
  createdDateTime?: string;
  completedDateTime?: string | null;
  assignments?: PlannerAssignments;
  hasDescription?: boolean;
  checklistItemCount?: number;
  activeChecklistItemCount?: number;
}

export interface PlannerChecklistItems {
  [itemId: string]: {
    "@odata.type"?: string;
    title?: string;
    isChecked?: boolean;
    orderHint?: string;
  } | null;
}

export interface PlannerTaskDetails {
  "@odata.etag"?: string;
  id: string;
  description?: string;
  checklist?: PlannerChecklistItems;
}

/** Human label for Planner's numeric priority. */
export function priorityLabel(priority: number | undefined): string {
  if (priority === undefined) return "unset";
  if (priority <= 1) return "urgent";
  if (priority <= 4) return "important";
  if (priority <= 7) return "medium";
  return "low";
}

export const PRIORITY_VALUES = { urgent: 1, important: 3, medium: 5, low: 9 } as const;
export type PriorityName = keyof typeof PRIORITY_VALUES;

export function progressLabel(percentComplete: number | undefined): string {
  if (percentComplete === 100) return "completed";
  if (percentComplete === 50) return "in progress";
  return "not started";
}
