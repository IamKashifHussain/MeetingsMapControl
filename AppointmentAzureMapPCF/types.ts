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
  ownerId?: string; // Added for filtering by current user
}