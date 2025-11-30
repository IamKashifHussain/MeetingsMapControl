import * as React from "react";
import * as atlas from "azure-maps-control";
import "azure-maps-control/dist/atlas.min.css";
import { IInputs } from "./generated/ManifestTypes";
import { AppointmentRecord } from "./types";

interface MapComponentProps {
  appointments: AppointmentRecord[];
  azureMapsKey: string;
  context: ComponentFramework.Context<IInputs>;
  currentUserName: string;
  totalAppointments: number;
  filteredAppointments: number;
}

type GeocodeCache = Record<string, atlas.data.Position | null>;

interface MarkerInfo {
  marker: atlas.HtmlMarker;
  appointment: AppointmentRecord;
}

const MapComponent: React.FC<MapComponentProps> = ({ 
  appointments, 
  azureMapsKey, 
  context,
  currentUserName,
  totalAppointments,
  filteredAppointments 
}) => {
  const mapRef = React.useRef<HTMLDivElement>(null);
  const mapInstanceRef = React.useRef<atlas.Map | null>(null);
  const popupRef = React.useRef<atlas.Popup | null>(null);
  const geocodeCacheRef = React.useRef<GeocodeCache>({});
  const markersRef = React.useRef<MarkerInfo[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [errorMessage, setErrorMessage] = React.useState<string>("");

  React.useEffect(() => {
    if (!mapRef.current) {
      setErrorMessage("Map container not found");
      setIsLoading(false);
      return;
    }

    if (!azureMapsKey) {
      setErrorMessage("Azure Maps subscription key is required");
      setIsLoading(false);
      return;
    }

    initializeMap();

    return () => {
      cleanup();
    };
  }, [azureMapsKey]);

  React.useEffect(() => {
    if (mapInstanceRef.current && popupRef.current) {
      updateMarkers();
    }
  }, [appointments]);

  const cleanup = () => {
    if (popupRef.current) {
      popupRef.current.close();
      popupRef.current = null;
    }
    
    if (markersRef.current.length > 0) {
      markersRef.current.forEach(markerInfo => {
        if (mapInstanceRef.current) {
          mapInstanceRef.current.markers.remove(markerInfo.marker);
        }
      });
      markersRef.current = [];
    }

    if (mapInstanceRef.current) {
      mapInstanceRef.current.dispose();
      mapInstanceRef.current = null;
    }
  };

  const initializeMap = () => {
    if (!mapRef.current || !azureMapsKey) return;

    try {
      const map = new atlas.Map(mapRef.current, {
        center: [-98.5795, 39.8283],
        zoom: 4,
        view: "Auto",
        authOptions: {
          authType: atlas.AuthenticationType.subscriptionKey,
          subscriptionKey: azureMapsKey,
        },
        style: "road",
        language: "en-US",
        showFeedbackLink: false,
        showLogo: false,
      });

      mapInstanceRef.current = map;

      const popup = new atlas.Popup({
        pixelOffset: [0, -18],
        closeButton: true,
      });
      popupRef.current = popup;

      map.events.add("ready", async () => {
        map.controls.add(
          [
            new atlas.control.ZoomControl(),
            new atlas.control.CompassControl(),
            new atlas.control.PitchControl(),
            new atlas.control.StyleControl({
              mapStyles: ["road", "satellite", "satellite_road_labels", "night", "road_shaded_relief"]
            }),
          ],
          {
            position: atlas.ControlPosition.TopRight,
          }
        );

        await updateMarkers();
        setIsLoading(false);
      });

      map.events.add("error", (error) => {
        console.error("Map error:", error);
        setErrorMessage("Failed to load map");
        setIsLoading(false);
      });

    } catch (error) {
      console.error("Map initialization error:", error);
      setErrorMessage("Failed to initialize map");
      setIsLoading(false);
    }
  };

  const geocodeAddress = async (address: string): Promise<atlas.data.Position | null> => {
    if (!address || !address.trim()) return null;

    const normalizedAddress = address.trim().toLowerCase();

    if (geocodeCacheRef.current[normalizedAddress] !== undefined) {
      return geocodeCacheRef.current[normalizedAddress];
    }

    try {
      const response = await fetch(
        `https://atlas.microsoft.com/search/address/json?api-version=1.0&subscription-key=${azureMapsKey}&query=${encodeURIComponent(
          address
        )}&limit=1`
      );

      if (!response.ok) {
        throw new Error(`Geocoding failed: ${response.statusText}`);
      }

      const data = await response.json();

      if (data.results && data.results.length > 0) {
        const position = new atlas.data.Position(
          data.results[0].position.lon,
          data.results[0].position.lat
        );
        geocodeCacheRef.current[normalizedAddress] = position;
        return position;
      }

      geocodeCacheRef.current[normalizedAddress] = null;
      return null;
    } catch (error) {
      console.error("Geocoding error for address:", address, error);
      geocodeCacheRef.current[normalizedAddress] = null;
      return null;
    }
  };

  const fetchRegardingAddress = async (
    regardingobjectid: ComponentFramework.EntityReference
  ): Promise<string | null> => {
    if (!regardingobjectid || !regardingobjectid.id || !regardingobjectid.etn) {
      return null;
    }

    try {
      const entityType = regardingobjectid.etn.toLowerCase();
      // Handle both string and object formats for entity ID
      const entityId = typeof regardingobjectid.id === 'string' 
        ? regardingobjectid.id 
        : regardingobjectid.id.guid;

      let selectQuery = "";
      if (entityType === "contact" || entityType === "account") {
        selectQuery = "?$select=address1_composite,address1_line1,address1_city,address1_stateorprovince,address1_postalcode,address1_country";
      } else if (entityType === "lead") {
        selectQuery = "?$select=address1_composite,address1_line1,address1_city,address1_stateorprovince,address1_postalcode,address1_country";
      } else {
        return null;
      }

      const record = await context.webAPI.retrieveRecord(entityType, entityId, selectQuery);

      if (record.address1_composite) {
        return record.address1_composite;
      }

      const addressParts = [
        record.address1_line1,
        record.address1_city,
        record.address1_stateorprovince,
        record.address1_postalcode,
        record.address1_country,
      ].filter(Boolean);

      return addressParts.length > 0 ? addressParts.join(", ") : null;
    } catch (error) {
      console.error("Error fetching regarding address:", error);
      return null;
    }
  };

  const formatDateTime = (date: Date): string => {
    if (!date) return "Not specified";
    
    try {
      return new Intl.DateTimeFormat("en-US", {
        dateStyle: "medium",
        timeStyle: "short",
      }).format(new Date(date));
    } catch {
      return date.toString();
    }
  };

  const calculateDuration = (start: Date, end: Date): string => {
    try {
      const startDate = new Date(start);
      const endDate = new Date(end);
      const durationMs = endDate.getTime() - startDate.getTime();
      const hours = Math.floor(durationMs / (1000 * 60 * 60));
      const minutes = Math.floor((durationMs % (1000 * 60 * 60)) / (1000 * 60));
      
      if (hours > 0) {
        return `${hours}h ${minutes}m`;
      }
      return `${minutes}m`;
    } catch {
      return "Unknown";
    }
  };

  const createPopupContent = (appt: AppointmentRecord): string => {
    const startTime = formatDateTime(appt.scheduledstart);
    const endTime = formatDateTime(appt.scheduledend);
    const duration = calculateDuration(appt.scheduledstart, appt.scheduledend);
    const description = appt.description || "No description available";
    const location = appt.location || "Location not specified";
    const regarding = appt.regardingobjectidname || "";

    return `
      <div style="padding: 15px; min-width: 280px; max-width: 350px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
        <h3 style="margin: 0 0 12px 0; color: #0078d4; font-size: 17px; font-weight: 600; border-bottom: 2px solid #0078d4; padding-bottom: 6px;">
          ${escapeHtml(appt.subject)}
        </h3>
        
        <div style="margin-bottom: 10px; background: #f5f5f5; padding: 8px; border-radius: 4px;">
          <div style="margin-bottom: 6px;">
            <span style="color: #333; font-weight: 600;">📅 Start:</span> 
            <span style="color: #555; margin-left: 5px;">${escapeHtml(startTime)}</span>
          </div>
          <div style="margin-bottom: 6px;">
            <span style="color: #333; font-weight: 600;">🕐 End:</span> 
            <span style="color: #555; margin-left: 5px;">${escapeHtml(endTime)}</span>
          </div>
          <div>
            <span style="color: #333; font-weight: 600;">⏱️ Duration:</span> 
            <span style="color: #555; margin-left: 5px;">${duration}</span>
          </div>
        </div>

        <div style="margin-bottom: 8px;">
          <span style="color: #333; font-weight: 600;">📍 Location:</span> 
          <div style="color: #555; margin-top: 4px; padding-left: 20px;">${escapeHtml(location)}</div>
        </div>

        ${
          regarding
            ? `
        <div style="margin-bottom: 8px;">
          <span style="color: #333; font-weight: 600;">👤 Regarding:</span> 
          <div style="color: #555; margin-top: 4px; padding-left: 20px;">${escapeHtml(regarding)}</div>
        </div>
        `
            : ""
        }

        ${
          description && description !== "No description available"
            ? `
        <div style="margin-top: 12px; padding-top: 10px; border-top: 1px solid #e0e0e0;">
          <span style="color: #333; font-weight: 600; display: block; margin-bottom: 6px;">📝 Description:</span>
          <div style="color: #555; max-height: 120px; overflow-y: auto; padding: 8px; background: #f9f9f9; border-radius: 4px; font-size: 13px; line-height: 1.5;">
            ${escapeHtml(description)}
          </div>
        </div>
        `
            : ""
        }
      </div>
    `;
  };

  const escapeHtml = (text: string): string => {
    if (!text) return "";
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
  };

  const updateMarkers = async () => {
    const map = mapInstanceRef.current;
    const popup = popupRef.current;

    if (!map || !popup) return;

    markersRef.current.forEach(markerInfo => {
      map.markers.remove(markerInfo.marker);
    });
    markersRef.current = [];

    const positions: atlas.data.Position[] = [];
    let markerCount = 0;

    for (const appt of appointments) {
      let address: string | null | undefined = appt.location;

      if (!address && appt.regardingobjectid) {
        address = await fetchRegardingAddress(appt.regardingobjectid);
      }

      if (!address) {
        console.warn(`No address found for appointment: ${appt.subject}`);
        continue;
      }

      const position = await geocodeAddress(address);

      if (!position) {
        console.warn(`Could not geocode address for appointment: ${appt.subject}`);
        continue;
      }

      positions.push(position);

      const marker = new atlas.HtmlMarker({
        color: "DodgerBlue",
        text: (markerCount + 1).toString(),
        position: position,
      });

      map.events.add("click", marker, () => {
        popup.setOptions({
          content: createPopupContent(appt),
          position: position,
        });
        popup.open(map);
      });

      map.markers.add(marker);
      markersRef.current.push({ marker, appointment: appt });
      markerCount++;
    }

    if (positions.length > 0) {
      const bounds = atlas.data.BoundingBox.fromPositions(positions);
      map.setCamera({
        bounds: bounds,
        padding: 80,
        maxZoom: 15,
      });
    } else {
      console.warn("No appointments could be mapped");
    }
  };

  return (
    <div style={{ width: "100%", height: "100%", position: "relative", minHeight: "500px" }}>
      {/* User info banner */}
      <div
        style={{
          position: "absolute",
          top: "10px",
          left: "10px",
          background: "rgba(255, 255, 255, 0.95)",
          padding: "12px 16px",
          borderRadius: "6px",
          boxShadow: "0 2px 8px rgba(0,0,0,0.15)",
          zIndex: 1000,
          fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
        }}
      >
        <div style={{ fontSize: "12px", color: "#666", marginBottom: "4px" }}>
          Showing appointments for
        </div>
        <div style={{ fontSize: "14px", fontWeight: "600", color: "#0078d4" }}>
          👤 {currentUserName}
        </div>
        <div style={{ fontSize: "11px", color: "#888", marginTop: "4px" }}>
          {filteredAppointments} of {totalAppointments} total appointments
        </div>
      </div>

      <div ref={mapRef} style={{ width: "100%", height: "100%" }} />
      
      {isLoading && (
        <div
          style={{
            position: "absolute",
            top: "50%",
            left: "50%",
            transform: "translate(-50%, -50%)",
            background: "rgba(255, 255, 255, 0.95)",
            padding: "25px",
            borderRadius: "8px",
            boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
            textAlign: "center",
            zIndex: 1000,
          }}
        >
          <div
            style={{
              border: "4px solid #f3f3f3",
              borderTop: "4px solid #0078d4",
              borderRadius: "50%",
              width: "50px",
              height: "50px",
              animation: "spin 1s linear infinite",
              margin: "0 auto 15px",
            }}
          />
          <p style={{ margin: 0, color: "#333", fontSize: "14px" }}>Loading your appointments...</p>
        </div>
      )}

      {errorMessage && (
        <div
          style={{
            position: "absolute",
            top: "50%",
            left: "50%",
            transform: "translate(-50%, -50%)",
            background: "#fff",
            padding: "25px",
            borderRadius: "8px",
            boxShadow: "0 4px 12px rgba(0,0,0,0.2)",
            textAlign: "center",
            maxWidth: "400px",
            zIndex: 1000,
          }}
        >
          <h3 style={{ color: "#d13438", marginTop: 0, marginBottom: "10px", fontSize: "18px" }}>
            ⚠️ Configuration Required
          </h3>
          <p style={{ color: "#555", margin: 0, fontSize: "14px", lineHeight: "1.5" }}>
            {errorMessage}
          </p>
          <p style={{ color: "#777", margin: "15px 0 0 0", fontSize: "12px" }}>
            Please configure your Azure Maps subscription key in the control properties.
          </p>
        </div>
      )}

      {!isLoading && !errorMessage && filteredAppointments === 0 && (
        <div
          style={{
            position: "absolute",
            top: "50%",
            left: "50%",
            transform: "translate(-50%, -50%)",
            background: "rgba(255, 255, 255, 0.95)",
            padding: "30px",
            borderRadius: "8px",
            boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
            textAlign: "center",
            maxWidth: "400px",
            zIndex: 1000,
          }}
        >
          <div style={{ fontSize: "48px", marginBottom: "15px" }}>📅</div>
          <h3 style={{ color: "#333", marginTop: 0, marginBottom: "10px", fontSize: "18px" }}>
            No Appointments Found
          </h3>
          <p style={{ color: "#666", margin: 0, fontSize: "14px", lineHeight: "1.5" }}>
            You don't have any appointments to display on the map.
          </p>
          {totalAppointments > 0 && (
            <p style={{ color: "#888", margin: "10px 0 0 0", fontSize: "12px" }}>
              ({totalAppointments} total appointments exist, but none are assigned to you)
            </p>
          )}
        </div>
      )}
    </div>
  );
};

export default MapComponent;