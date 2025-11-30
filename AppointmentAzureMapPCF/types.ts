// Shared type definitions for the PCF control

export interface AppointmentRecord {
  id: string;
  subject: string;
  scheduledstart: Date;
  scheduledend: Date;
  location?: string;
  description?: string;
  regardingobjectid?: ComponentFramework.EntityReference;
  regardingobjectidname?: string;
  ownerId?: string;
}

export type DueFilter = 
  | "all"
  | "overdue"
  | "today"
  | "tomorrow"
  | "next7days"
  | "next30days"
  | "next90days"
  | "next6months"
  | "next12months";

export interface FilterOptions {
  dueFilter: DueFilter;
  searchText?: string;
}