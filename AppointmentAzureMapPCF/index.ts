import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import MapComponent from "./MapComponent";
import { AppointmentRecord, FilterOptions } from "./types";

interface XrmGlobal {
  Xrm?: {
    Utility?: {
      getGlobalContext: () => {
        userSettings: { userId?: string };
      };
    };
  };
}

export class AppointmentAzureMapPCF
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private container: HTMLDivElement;
  private root: Root | null = null;
  private notifyOutputChanged: () => void;
  private azureMapsKey = "";
  private isLoadingConfig = true;
  private allAppointments: AppointmentRecord[] = [];
  private filteredAppointments: AppointmentRecord[] = [];
  private isLoadingAppointments = true;
  private currentFilter: FilterOptions = {
    dueFilter: "today",
  };

  private context: ComponentFramework.Context<IInputs>;
  private currentUserAddress = "";
  private currentUserId = "";
  private currentUserName = "";

  private showRoute = false;
  private refreshTrigger = 0;

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.container = container;
    this.notifyOutputChanged = notifyOutputChanged;
    this.context = context;

    this.root = createRoot(this.container);

    void this.initializeData(context);
  }

  private async initializeData(
    context: ComponentFramework.Context<IInputs>
  ): Promise<void> {
    try {
      this.getUserId(context);

      if (!this.currentUserId) {
        this.isLoadingConfig = false;
        this.isLoadingAppointments = false;
        this.renderComponent();
        return;
      }

      await Promise.all([
        this.fetchAzureMapsKeyFromConfig(context),
        this.fetchCurrentUserAddress(context),
      ]);

      await this.fetchUserAppointments(context);

      this.applyFilters();
    } finally {
      this.renderComponent();
    }
  }

  private getUserId(context: ComponentFramework.Context<IInputs>): void {
    let userId: string | null = null;

    try {
      if (context.parameters?.userId?.raw) {
        userId = context.parameters.userId.raw;
      } else if (context.userSettings?.userId) {
        userId = context.userSettings.userId;
      } else {
        try {
          const globalObj = window as unknown as XrmGlobal;
          if (globalObj.Xrm?.Utility?.getGlobalContext) {
            const us = globalObj.Xrm.Utility.getGlobalContext().userSettings;
            if (us?.userId) userId = us.userId;
          }
        } catch {
          // Xrm not available
        }
      }

      if (userId) {
        if (userId.includes("@")) {
          this.currentUserId = userId.toLowerCase();
        } else {
          this.currentUserId = userId.replace(/[{}]/g, "").toLowerCase();
        }
      } else {
        this.currentUserId = "";
      }
    } catch {
      this.currentUserId = "";
    }
  }

  private async fetchAzureMapsKeyFromConfig(
    context: ComponentFramework.Context<IInputs>
  ): Promise<void> {
    try {
      const result = await context.webAPI.retrieveMultipleRecords(
        "ti_mapconfiguration",
        "?$select=ti_azuremapskey&$filter=statecode eq 0&$top=1"
      );

      if (result.entities.length > 0) {
        this.azureMapsKey = result.entities[0].ti_azuremapskey ?? "";
      }
    } finally {
      this.isLoadingConfig = false;
    }
  }

  private async fetchCurrentUserAddress(
    context: ComponentFramework.Context<IInputs>
  ): Promise<void> {
    if (!this.currentUserId) return;

    try {
      try {
        const user = await context.webAPI.retrieveRecord(
          "systemuser",
          this.currentUserId,
          "?$select=address1_composite,fullname"
        );

        this.currentUserAddress = user.address1_composite ?? "";
        this.currentUserName = user.fullname ?? "Current User";
        return;
      } catch {
        // retrieveRecord failed
      }

      const result = await context.webAPI.retrieveMultipleRecords(
        "systemuser",
        `?$select=address1_composite,fullname&$filter=systemuserid eq ${this.currentUserId}&$top=1`
      );

      if (result.entities.length > 0) {
        this.currentUserAddress = result.entities[0].address1_composite ?? "";
        this.currentUserName = result.entities[0].fullname ?? "Current User";
      }
    } catch {
      this.currentUserAddress = "";
      this.currentUserName = "Current User";
    }
  }

  private async fetchUserAppointments(
    context: ComponentFramework.Context<IInputs>
  ): Promise<void> {
    if (!this.currentUserId) {
      this.allAppointments = [];
      this.isLoadingAppointments = false;
      return;
    }

    try {
      const query =
        `?$select=subject,scheduledstart,scheduledend,location,description,activityid,_regardingobjectid_value` +
        `&$filter=_ownerid_value eq ${this.currentUserId} and isonlinemeeting eq false and (statecode eq 0 or statecode eq 3)` +
        `&$orderby=scheduledstart asc&$top=5000`;

      const result = await context.webAPI.retrieveMultipleRecords(
        "appointment",
        query
      );

      this.allAppointments = result.entities.map((entity) => {
        let regardingRef: ComponentFramework.EntityReference | undefined;

        if (entity._regardingobjectid_value) {
          regardingRef = {
            id: { guid: entity._regardingobjectid_value },
            etn:
              entity[
                "_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname"
              ] ?? "",
            name:
              entity[
                "_regardingobjectid_value@OData.Community.Display.V1.FormattedValue"
              ] ?? "",
          };
        }

        return {
          id: entity.activityid,
          subject: entity.subject ?? "No Subject",
          scheduledstart: new Date(entity.scheduledstart),
          scheduledend: new Date(entity.scheduledend),
          location: entity.location ?? "",
          description: entity.description ?? "",
          regardingobjectid: regardingRef,
          regardingobjectidname:
            entity[
              "_regardingobjectid_value@OData.Community.Display.V1.FormattedValue"
            ] ?? "",
          ownerId: this.currentUserId,
        };
      });
    } finally {
      this.isLoadingAppointments = false;
    }
  }

  private applyFilters(): void {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);

    let filtered = [...this.allAppointments];

    switch (this.currentFilter.dueFilter) {
      case "overdue":
        filtered = filtered.filter((a) => a.scheduledstart < today);
        break;
      case "tomorrow":
        filtered = filtered.filter(
          (a) =>
            a.scheduledstart >= tomorrow &&
            a.scheduledstart <
              new Date(tomorrow.getTime() + 24 * 60 * 60 * 1000)
        );
        break;
      case "next7days":
        filtered = filtered.filter(
          (a) =>
            a.scheduledstart >= tomorrow &&
            a.scheduledstart <
              new Date(tomorrow.getTime() + 7 * 24 * 60 * 60 * 1000)
        );
        break;
      case "next30days":
        filtered = filtered.filter(
          (a) =>
            a.scheduledstart >= tomorrow &&
            a.scheduledstart <
              new Date(tomorrow.getTime() + 30 * 24 * 60 * 60 * 1000)
        );
        break;
      case "next90days":
        filtered = filtered.filter(
          (a) =>
            a.scheduledstart >= tomorrow &&
            a.scheduledstart <
              new Date(tomorrow.getTime() + 90 * 24 * 60 * 60 * 1000)
        );
        break;
      case "next6months": {
        const d = new Date(tomorrow);
        d.setMonth(d.getMonth() + 6);
        filtered = filtered.filter(
          (a) => a.scheduledstart >= tomorrow && a.scheduledstart < d
        );
        break;
      }
      case "next12months": {
        const d = new Date(tomorrow);
        d.setMonth(d.getMonth() + 12);
        filtered = filtered.filter(
          (a) => a.scheduledstart >= tomorrow && a.scheduledstart < d
        );
        break;
      }
      case "customDateRange": {
        if (this.currentFilter.customDateRange) {
          const { startDate, endDate } = this.currentFilter.customDateRange;

          const rangeStart = new Date(
            startDate.getUTCFullYear(),
            startDate.getUTCMonth(),
            startDate.getUTCDate(),
            0,
            0,
            0,
            0
          );

          const rangeEnd = new Date(
            endDate.getUTCFullYear(),
            endDate.getUTCMonth(),
            endDate.getUTCDate(),
            23,
            59,
            59,
            999
          );

          filtered = filtered.filter(
            (a) =>
              a.scheduledstart >= rangeStart &&
              a.scheduledstart <= rangeEnd
          );
        }
        break;
      }
      case "all":
        break;
      case "today":
      default:
        filtered = filtered.filter(
          (a) =>
            a.scheduledstart >= today &&
            a.scheduledstart <
              new Date(today.getTime() + 24 * 60 * 60 * 1000)
        );
        break;
    }

    this.filteredAppointments = filtered;
  }

  private handleFilterChange = (newFilter: FilterOptions): void => {
    this.currentFilter = newFilter;
    this.applyFilters();
    this.renderComponent();
  };

  private handleRouteToggle = (enabled: boolean): void => {
    this.showRoute = enabled;
    this.renderComponent();
  };

  private handleRefresh = async (): Promise<void> => {
    this.isLoadingAppointments = true;
    this.renderComponent();

    this.getUserId(this.context);

    await Promise.all([
      this.fetchUserAppointments(this.context),
      this.fetchCurrentUserAddress(this.context),
    ]);

    this.applyFilters();
    this.refreshTrigger++;
    this.renderComponent();
  };

  private renderComponent(): void {
    if (!this.root) return;

    if (this.isLoadingConfig || this.isLoadingAppointments) {
      this.root.render(
        React.createElement(
          "div",
          { style: { padding: "20px", textAlign: "center" } },
          this.isLoadingConfig
            ? "Loading configuration..."
            : "Loading appointments..."
        )
      );
      return;
    }

    if (!this.azureMapsKey || !this.currentUserId) return;

    this.root.render(
      React.createElement(MapComponent, {
        key: this.refreshTrigger,
        appointments: this.filteredAppointments,
        allAppointmentsCount: this.allAppointments.length,
        azureMapsKey: this.azureMapsKey,
        context: this.context,
        currentUserName: this.currentUserName,
        currentUserAddress: this.currentUserAddress,
        currentFilter: this.currentFilter,
        showRoute: this.showRoute,
        onFilterChange: this.handleFilterChange,
        onRouteToggle: this.handleRouteToggle,
        onRefresh: this.handleRefresh,
      })
    );
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this.context = context;

    const previous = this.currentUserId;
    this.getUserId(context);

    if (previous !== this.currentUserId && this.currentUserId) {
      void this.initializeData(context);
    }
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    if (this.root) {
      this.root.unmount();
      this.root = null;
    }
  }
}