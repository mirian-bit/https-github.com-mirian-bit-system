/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export enum LeaveType {
  None = "",
  Self = "自休",
  System = "系休",
  National = "國定"
}

export interface SkillRequirement {
  skill: string;
  count: number;
}

export interface GroupConfig {
  id: string;
  name: string;
  maxLeavePerDay: number;
  skillRequirements: SkillRequirement[];
}

export interface Department {
  id: string;
  name: string;
  skills: string[];
  groups: GroupConfig[];
  attendanceTimes?: string[]; // 該課別專屬的出勤時間選單
}

export interface Employee {
  id: string;
  name: string;
  deptId: string;
  groupId: string;
  skills: string[];
  nationalHolidayQuota: number; // 本週期可排國定天數
  systemLeaveQuota: number; // 本週期應排系休天數
}

export interface ShiftRecord {
  employeeId: string;
  date: string; // ISO string YYYY-MM-DD
  type: "shift" | "leave";
  code: string; // e.g. "900", "900(技能)", "自休", etc.
  time?: string; // e.g. "08:00-17:00"
  isManual?: boolean; // 標記是否為手動修改，避免排班引擎覆蓋
}

export interface DailyRequirement {
  id: string; // groupId_date
  groupId: string;
  date: string;
  maxLeave: number;
  skillRequirements: SkillRequirement[];
}

export interface ScheduleState {
  departments: Department[];
  employees: Employee[];
  records: Record<string, ShiftRecord>; // key: employeeId_date
  dailyRequirements: Record<string, DailyRequirement>; // key: groupId_date
  currentPeriodStart: string; // ISO string YYYY-MM-DD
  nationalHolidays: string[]; // 國定假日日期清單 (YYYY-MM-DD)
  attendanceTimes: string[]; // 出勤時間選單
  leaveInterval: {
    start: string;
    end: string;
  };
}
