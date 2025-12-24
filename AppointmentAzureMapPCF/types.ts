import { data } from "azure-maps-control";

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
  | "next12months"
  | "customDateRange";

export interface DateRange {
  startDate: Date;
  endDate: Date;
}

export interface FilterOptions {
  dueFilter: DueFilter;
  customDateRange?: DateRange;
}

export interface UserLocation {
  address: string;
  position: data.Position;
}