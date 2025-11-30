import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import MapComponent from "./MapComponent";
import { AppointmentRecord } from "./types";

export class AppointmentAzureMapPCF
  implements ComponentFramework.ReactControl<IInputs, IOutputs>
{
  private notifyOutputChanged!: () => void;
  private azureMapsKey = "";
  private isLoadingConfig = true;

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary
  ): void {
    this.notifyOutputChanged = notifyOutputChanged;
    void this.fetchAzureMapsKeyFromConfig(context);
  }

  private async fetchAzureMapsKeyFromConfig(
    context: ComponentFramework.Context<IInputs>
  ): Promise<void> {
    try {
      console.log("Fetching Azure Maps key from ti_mapconfiguration...");

      const result = await context.webAPI.retrieveMultipleRecords(
        "ti_mapconfiguration",
        "?$select=ti_azuremapskey&$top=1"
      );

      if (result.entities.length > 0) {
        const configRecord = result.entities[0];
        this.azureMapsKey = configRecord.ti_azuremapskey ?? "";

        if (this.azureMapsKey) {
          console.log("✅ Azure Maps key successfully loaded");
        } else {
          console.warn(
            "⚠️ ti_mapconfiguration record found, but ti_azuremapskey field is empty"
          );
        }
      } else {
        console.warn("⚠️ No records found in ti_mapconfiguration table");
      }
    } catch (error) {
      console.error(
        "❌ Error fetching Azure Maps key from ti_mapconfiguration:",
        error
      );
    } finally {
      this.isLoadingConfig = false;
      this.notifyOutputChanged();
    }
  }

  public updateView(
    context: ComponentFramework.Context<IInputs>
  ): React.ReactElement {
    const dataset = context.parameters.appointments;

    if (this.isLoadingConfig) {
      return React.createElement(
        "div",
        {
          style: {
            padding: "20px",
            textAlign: "center",
            fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
          },
        },
        "Loading map configuration..."
      );
    }

    if (!this.azureMapsKey) {
      return React.createElement(
        "div",
        {
          style: {
            padding: "30px",
            textAlign: "center",
            fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
            color: "#d13438",
          },
        },
        [
          React.createElement(
            "h3",
            { key: "title", style: { marginTop: 0 } },
            "⚠️ Configuration Error"
          ),
          React.createElement(
            "p",
            { key: "message", style: { color: "#555" } },
            "Azure Maps key not found in ti_mapconfiguration table."
          ),
          React.createElement(
            "p",
            { key: "help", style: { fontSize: "12px", color: "#777" } },
            "Please ensure a record exists in ti_mapconfiguration with ti_azuremapskey field populated."
          ),
        ]
      );
    }

    if (!dataset || !dataset.sortedRecordIds || dataset.loading) {
      return React.createElement(
        "div",
        { style: { padding: "20px", textAlign: "center" } },
        "Loading appointments..."
      );
    }

    const currentUserId = context.userSettings.userId;

    const allAppointments: AppointmentRecord[] = dataset.sortedRecordIds.map(
      (id) => {
        const record = dataset.records[id];

        // Regarding object
        const regardingValue = record.getValue("regardingobjectid");
        const regardingRef =
          regardingValue && typeof regardingValue === "object" && "id" in regardingValue
            ? (regardingValue as ComponentFramework.EntityReference)
            : undefined;

        // Owner object
        const ownerValue = record.getValue("ownerid");
        let ownerId = "";

        if (ownerValue && typeof ownerValue === "object" && "id" in ownerValue) {
          const ownerRef = ownerValue as
            | ComponentFramework.EntityReference
            | { id: { guid: string } };

          const idValue = ownerRef.id;

          // Store ID as-is without converting to lowercase
          if (typeof idValue === "string") {
            ownerId = idValue;
          } else if (
            idValue &&
            typeof idValue === "object" &&
            "guid" in idValue &&
            typeof idValue.guid === "string"
          ) {
            ownerId = idValue.guid;
          }
        }

        return {
          id: record.getRecordId(),
          subject:
            record.getFormattedValue("subject") ||
            (record.getValue("subject") as string) ||
            "No Subject",
          scheduledstart: record.getValue("scheduledstart") as Date,
          scheduledend: record.getValue("scheduledend") as Date,
          location: (record.getValue("location") as string) ?? "",
          description: (record.getValue("description") as string) ?? "",
          regardingobjectid: regardingRef,
          regardingobjectidname: record.getFormattedValue("regardingobjectid"),
          ownerId,
        };
      }
    );

    const userAppointments = allAppointments.filter(
      (appt) => appt.ownerId === currentUserId
    );

    return React.createElement(MapComponent, {
      appointments: userAppointments,
      azureMapsKey: this.azureMapsKey,
      context,
      currentUserName: context.userSettings.userName ?? "Current User",
      totalAppointments: allAppointments.length,
      filteredAppointments: userAppointments.length,
    });
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    // Cleanup if needed
  }
}
