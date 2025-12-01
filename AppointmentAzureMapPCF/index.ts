import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import MapComponent from "./MapComponent";
import { AppointmentRecord, FilterOptions } from "./types";

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
    dueFilter: "all",
    searchText: ""
  };
  private context: ComponentFramework.Context<IInputs>;

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.container = container;
    this.notifyOutputChanged = notifyOutputChanged;
    this.context = context;

    // Create React root
    this.root = createRoot(this.container);

    void this.initializeData(context);
  }

  private async initializeData(context: ComponentFramework.Context<IInputs>): Promise<void> {
    try {
      await this.fetchAzureMapsKeyFromConfig(context);
      await this.fetchUserAppointments(context);
      this.applyFilters();
    } catch (error) {
      console.error("Error initializing data:", error);
    } finally {
      this.renderComponent();
    }
  }

  private async fetchAzureMapsKeyFromConfig(
    context: ComponentFramework.Context<IInputs>
  ): Promise<void> {
    try {
      const result = await context.webAPI.retrieveMultipleRecords(
        "ti_mapconfiguration",
        "?$select=ti_azuremapskey&$top=1"
      );

      if (result.entities.length > 0) {
        const configRecord = result.entities[0];
        this.azureMapsKey = configRecord.ti_azuremapskey ?? "";
        if (!this.azureMapsKey) {
          console.warn("Azure Maps key is empty in configuration record.");
        } else {
          console.log("✅ Azure Maps key loaded successfully");
        }
      } else {
        console.warn("No records found in ti_mapconfiguration table.");
      }
    } catch (error) {
      console.error("Error fetching Azure Maps key:", error);
    } finally {
      this.isLoadingConfig = false;
    }
  }

  private async fetchUserAppointments(
    context: ComponentFramework.Context<IInputs>
  ): Promise<void> {
    try {
      const currentUserId = context.userSettings.userId;

      // ✅ FIXED: Includes regardingobjectid
      const fetchXml = `
        <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
          <entity name="appointment">
            <attribute name="subject" />
            <attribute name="scheduledstart" />
            <attribute name="scheduledend" />
            <attribute name="location" />
            <attribute name="description" />
            <attribute name="activityid" />
            <attribute name="regardingobjectid" />
            <order attribute="scheduledstart" descending="false" />
            <filter type="and">
              <condition attribute="ownerid" operator="eq" value="${currentUserId}" />
              <condition attribute="statecode" operator="eq" value="0" />
            </filter>
          </entity>
        </fetch>
      `;

      const encodedFetchXml = encodeURIComponent(fetchXml);

      const result = await context.webAPI.retrieveMultipleRecords(
        "appointment",
        `?fetchXml=${encodedFetchXml}`
      );

      console.log(`📊 Fetched ${result.entities.length} appointments`);

      this.allAppointments = result.entities.map((entity) => {
        let regardingRef: ComponentFramework.EntityReference | undefined;

        // ✅ FIXED: Proper EntityReference construction with correct property name
        if (entity._regardingobjectid_value) {
          // The property name for lookup logical name in FetchXML results
          const regardingType =
            entity["_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname"] ||
            entity["regardingobjectid@Microsoft.Dynamics.CRM.lookuplogicalname"];
          
          const regardingName =
            entity["_regardingobjectid_value@OData.Community.Display.V1.FormattedValue"];

          console.log(
            `   📌 Appointment "${entity.subject}": Regarding Type = "${regardingType}", ID = ${entity._regardingobjectid_value}, Name = "${regardingName}"`
          );

          // ✅ FIXED: Create proper EntityReference with id as object containing guid
          regardingRef = {
            id: { guid: entity._regardingobjectid_value },
            etn: regardingType || "",
            name: regardingName || "",
          } as ComponentFramework.EntityReference;
        } else {
          console.warn(
            `   ⚠️  Appointment "${entity.subject}": NO regarding object found`
          );
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
            entity["_regardingobjectid_value@OData.Community.Display.V1.FormattedValue"],
          ownerId: currentUserId.toLowerCase()
        } as AppointmentRecord;
      });

      console.log(`✅ Successfully mapped ${this.allAppointments.length} appointments`);
    } catch (error) {
      console.error("Error fetching appointments:", error);
      this.allAppointments = [];
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
      case "overdue": {
        filtered = filtered.filter(appt => appt.scheduledstart < now);
        break;
      }
      case "today": {
        const endOfToday = new Date(today);
        endOfToday.setDate(endOfToday.getDate() + 1);
        filtered = filtered.filter(appt => appt.scheduledstart >= today && appt.scheduledstart < endOfToday);
        break;
      }
      case "tomorrow": {
        const endOfTomorrow = new Date(tomorrow);
        endOfTomorrow.setDate(endOfTomorrow.getDate() + 1);
        filtered = filtered.filter(appt => appt.scheduledstart >= tomorrow && appt.scheduledstart < endOfTomorrow);
        break;
      }
      case "next7days": {
        const next7Days = new Date(today);
        next7Days.setDate(next7Days.getDate() + 7);
        filtered = filtered.filter(appt => appt.scheduledstart >= today && appt.scheduledstart < next7Days);
        break;
      }
      case "next30days": {
        const next30Days = new Date(today);
        next30Days.setDate(next30Days.getDate() + 30);
        filtered = filtered.filter(appt => appt.scheduledstart >= today && appt.scheduledstart < next30Days);
        break;
      }
      case "next90days": {
        const next90Days = new Date(today);
        next90Days.setDate(next90Days.getDate() + 90);
        filtered = filtered.filter(appt => appt.scheduledstart >= today && appt.scheduledstart < next90Days);
        break;
      }
      case "next6months": {
        const next6Months = new Date(today);
        next6Months.setMonth(next6Months.getMonth() + 6);
        filtered = filtered.filter(appt => appt.scheduledstart >= today && appt.scheduledstart < next6Months);
        break;
      }
      case "next12months": {
        const next12Months = new Date(today);
        next12Months.setMonth(next12Months.getMonth() + 12);
        filtered = filtered.filter(appt => appt.scheduledstart >= today && appt.scheduledstart < next12Months);
        break;
      }
      case "all":
      default:
        break;
    }

    if (this.currentFilter.searchText?.trim()) {
      const searchLower = this.currentFilter.searchText.toLowerCase().trim();
      filtered = filtered.filter(appt =>
        appt.subject.toLowerCase().includes(searchLower) ||
        appt.location?.toLowerCase().includes(searchLower) ||
        appt.description?.toLowerCase().includes(searchLower) ||
        appt.regardingobjectidname?.toLowerCase().includes(searchLower)
      );
    }

    this.filteredAppointments = filtered;
    console.log(`Filtered appointments: ${filtered.length} of ${this.allAppointments.length}`);
  }

  private handleFilterChange = (newFilter: FilterOptions): void => {
    this.currentFilter = newFilter;
    this.applyFilters();
    this.renderComponent();
  }

  private handleRefresh = async (): Promise<void> => {
    this.isLoadingAppointments = true;
    this.renderComponent();
    await this.fetchUserAppointments(this.context);
    this.applyFilters();
    this.renderComponent();
  }

  private renderComponent(): void {
    if (!this.root) return;

    if (this.isLoadingConfig || this.isLoadingAppointments) {
      this.root.render(
        React.createElement("div", { style: { padding: "20px", textAlign: "center", fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif" } }, "Loading appointments...")
      );
      return;
    }

    if (!this.azureMapsKey) {
      this.root.render(
        React.createElement("div", { style: { padding: "30px", textAlign: "center", fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif", color: "#d13438" } },
          [
            React.createElement("h3", { key: "title", style: { marginTop: 0 } }, "⚠️ Configuration Error"),
            React.createElement("p", { key: "message", style: { color: "#555" } }, "Azure Maps key not found in ti_mapconfiguration table."),
            React.createElement("p", { key: "help", style: { fontSize: "12px", color: "#777" } }, "Please ensure a record exists in ti_mapconfiguration with ti_azuremapskey field populated.")
          ]
        )
      );
      return;
    }

    this.root.render(
      React.createElement(MapComponent, {
        appointments: this.filteredAppointments,
        allAppointmentsCount: this.allAppointments.length,
        azureMapsKey: this.azureMapsKey,
        context: this.context,
        currentUserName: this.context.userSettings.userName ?? "Current User",
        currentFilter: this.currentFilter,
        onFilterChange: this.handleFilterChange,
        onRefresh: this.handleRefresh,
      })
    );
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this.context = context;
    // No re-render needed unless data changes
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