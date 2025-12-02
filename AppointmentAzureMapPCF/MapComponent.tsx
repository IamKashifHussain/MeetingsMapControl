// MapComponent.tsx
import * as React from "react";
import * as atlas from "azure-maps-control";
import "azure-maps-control/dist/atlas.min.css";
import { IInputs } from "./generated/ManifestTypes";
import {
  AppointmentRecord,
  DueFilter,
  FilterOptions,
  UserLocation,
} from "./types";

interface MapComponentProps {
  appointments: AppointmentRecord[];
  allAppointmentsCount: number;
  azureMapsKey: string;
  context: ComponentFramework.Context<IInputs>;
  currentUserName: string;
  currentUserAddress: string;
  currentFilter: FilterOptions;
  onFilterChange: (filter: FilterOptions) => void;
  onRefresh: () => void;
}

type GeocodeCache = Record<string, atlas.data.Position | null>;

interface MarkerInfo {
  marker: atlas.HtmlMarker;
  appointment: AppointmentRecord;
}

const MapComponent: React.FC<MapComponentProps> = ({
  appointments,
  allAppointmentsCount,
  azureMapsKey,
  context,
  currentUserName,
  currentUserAddress,
  currentFilter,
  onFilterChange,
  onRefresh,
}) => {
  const mapRef = React.useRef<HTMLDivElement>(null);
  const mapInstanceRef = React.useRef<atlas.Map | null>(null);
  const popupRef = React.useRef<atlas.Popup | null>(null);
  const geocodeCacheRef = React.useRef<GeocodeCache>({});
  const markersRef = React.useRef<MarkerInfo[]>([]);
  const userMarkerRef = React.useRef<atlas.HtmlMarker | null>(null);
  const abortControllerRef = React.useRef<AbortController | null>(null);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [errorMessage, setErrorMessage] = React.useState<string>("");
  const [searchText, setSearchText] = React.useState<string>(
    currentFilter.searchText || ""
  );
  const [userLocation, setUserLocation] = React.useState<UserLocation | null>(
    null
  );

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
    if (mapInstanceRef.current && popupRef.current && !isLoading) {
      abortControllerRef.current = new AbortController();
      void updateMarkers();
    }

    return () => {
      abortControllerRef.current?.abort();
    };
  }, [appointments]);

  React.useEffect(() => {
    if (currentUserAddress && mapInstanceRef.current && !isLoading) {
      void geocodeAndDisplayUserLocation();
    }
  }, [currentUserAddress, isLoading]);

  const cleanup = () => {
    abortControllerRef.current?.abort();

    if (popupRef.current) {
      popupRef.current.close();
      popupRef.current = null;
    }

    if (userMarkerRef.current && mapInstanceRef.current) {
      mapInstanceRef.current.markers.remove(userMarkerRef.current);
      userMarkerRef.current = null;
    }

    if (markersRef.current.length > 0) {
      markersRef.current.forEach((markerInfo) => {
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
              mapStyles: [
                "road",
                "satellite",
                "satellite_road_labels",
                "night",
                "road_shaded_relief",
              ],
            }),
          ],
          {
            position: atlas.ControlPosition.TopRight,
          }
        );

        await updateMarkers();
        if (currentUserAddress) {
          await geocodeAndDisplayUserLocation();
        }
        setIsLoading(false);
      });

      map.events.add("error", (error) => {
        if (markersRef.current.length === 0) {
          setErrorMessage("Failed to load map");
        }
        setIsLoading(false);
      });
    } catch (error) {
      setErrorMessage("Failed to initialize map");
      setIsLoading(false);
    }
  };

  const geocodeAddress = async (
    address: string
  ): Promise<atlas.data.Position | null> => {
    if (!address || !address.trim()) {
      return null;
    }

    const normalizedAddress = address.trim().toLowerCase();

    if (geocodeCacheRef.current[normalizedAddress] !== undefined) {
      return geocodeCacheRef.current[normalizedAddress];
    }

    try {
      const response = await fetch(
        `https://atlas.microsoft.com/search/address/json?api-version=1.0&subscription-key=${azureMapsKey}&query=${encodeURIComponent(
          address
        )}&limit=1`,
        { signal: abortControllerRef.current?.signal }
      );

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const data = await response.json();

      if (data.results && data.results.length > 0) {
        const result = data.results[0];

        if (
          result.position &&
          typeof result.position.lon === "number" &&
          typeof result.position.lat === "number"
        ) {
          const lon = result.position.lon;
          const lat = result.position.lat;
          const position: atlas.data.Position = [lon, lat];
          geocodeCacheRef.current[normalizedAddress] = position;
          return position;
        }
      }

      geocodeCacheRef.current[normalizedAddress] = null;
      return null;
    } catch (error) {
      if ((error as Error).name !== "AbortError") {
        geocodeCacheRef.current[normalizedAddress] = null;
      }
      return null;
    }
  };

  const geocodeAndDisplayUserLocation = async () => {
    const map = mapInstanceRef.current;
    const popup = popupRef.current;

    if (!map || !popup || !currentUserAddress) return;

    // Remove existing user marker if any
    if (userMarkerRef.current) {
      map.markers.remove(userMarkerRef.current);
      userMarkerRef.current = null;
    }

    const position = await geocodeAddress(currentUserAddress);

    if (position) {
      setUserLocation({ address: currentUserAddress, position });

      const userMarker = new atlas.HtmlMarker({
        color: "green",
        position: position,
      });

      map.events.add("click", userMarker, () => {
        popup.close();
        popup.setOptions({
          content: createUserPopupContent(currentUserName, currentUserAddress),
          position: position,
        });
        popup.open(map);
      });

      map.markers.add(userMarker);
      userMarkerRef.current = userMarker;

      map.events.add("click", userMarker, () => {
        popup.close();
        popup.setOptions({
          content: createUserPopupContent(currentUserName, currentUserAddress),
          position: position,
        });
        popup.open(map);
      });

      map.markers.add(userMarker);
      userMarkerRef.current = userMarker;
    }
  };

  const createUserPopupContent = (
    userName: string,
    address: string
  ): string => {
    return `
    <div style="
      padding: 16px;
      min-width: 260px;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background: rgba(32, 32, 32, 0.95);
      backdrop-filter: blur(8px);
      border-radius: 12px;
      box-shadow: 0 8px 20px rgba(0,0,0,0.4);
      color: #f1f1f1;
      transition: transform 0.2s ease-in-out;
    ">
      
      <div style="
        display: flex;
        align-items: center;
        margin-bottom: 12px;
      ">
        <svg xmlns="http://www.w3.org/2000/svg" style="width:20px;height:20px;margin-right:8px;" fill="#00d4ff" viewBox="0 0 24 24">
          <path d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5c-1.38 0-2.5-1.12-2.5-2.5S10.62 6.5 12 6.5s2.5 1.12 2.5 2.5S13.38 11.5 12 11.5z"/>
        </svg>
        <h3 style="margin:0; font-size: 16px; font-weight: 600; color:#00d4ff;">Your Location</h3>
      </div>
      
      <div style="margin-bottom: 10px;">
        <div style="font-size: 10px; font-weight: 600; color: #aaa; text-transform: uppercase; margin-bottom: 2px;">User</div>
        <div style="font-size: 14px; font-weight: 500; color: #fff;">${escapeHtml(
          userName
        )}</div>
      </div>
      
      <div>
        <div style="font-size: 10px; font-weight: 600; color: #aaa; text-transform: uppercase; margin-bottom: 2px;">Address</div>
        <div style="font-size: 13px; line-height: 1.4; color: #ccc;">${escapeHtml(
          address
        )}</div>
      </div> 
    </div>
  `;
  };

  const batchGeocodeAddresses = async (
    addresses: string[]
  ): Promise<Map<string, atlas.data.Position | null>> => {
    const results = new Map<string, atlas.data.Position | null>();
    const batchSize = 5;
    const delayMs = 100;

    for (let i = 0; i < addresses.length; i += batchSize) {
      if (abortControllerRef.current?.signal.aborted) {
        break;
      }

      const batch = addresses.slice(i, i + batchSize);

      await Promise.all(
        batch.map(async (address) => {
          const position = await geocodeAddress(address);
          results.set(address, position);
        })
      );

      if (i + batchSize < addresses.length) {
        await new Promise((resolve) => setTimeout(resolve, delayMs));
      }
    }

    return results;
  };

  const fetchRegardingAddress = async (
    regardingobjectid: ComponentFramework.EntityReference
  ): Promise<string | null> => {
    if (!regardingobjectid?.id || !regardingobjectid?.etn) {
      return null;
    }

    const entityType = regardingobjectid.etn.toLowerCase();

    let entityId: string;
    if (typeof regardingobjectid.id === "string") {
      entityId = regardingobjectid.id;
    } else if (
      regardingobjectid.id &&
      typeof regardingobjectid.id === "object" &&
      "guid" in regardingobjectid.id
    ) {
      entityId = (regardingobjectid.id as { guid: string }).guid;
    } else {
      return null;
    }

    try {
      if (["contact", "account"].includes(entityType)) {
        const record = await context.webAPI.retrieveRecord(
          entityType,
          entityId,
          "?$select=address1_composite"
        );
        return record.address1_composite ?? null;
      }

      if (entityType === "opportunity") {
        const opp = await context.webAPI.retrieveRecord(
          "opportunity",
          entityId,
          "?$select=_customerid_value&$expand=customerid_account($select=address1_composite),customerid_contact($select=address1_composite)"
        );

        if (opp.customerid_account?.address1_composite) {
          return opp.customerid_account.address1_composite;
        } else if (opp.customerid_contact?.address1_composite) {
          return opp.customerid_contact.address1_composite;
        }
        return null;
      }

      return null;
    } catch (error) {
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

  const createPopupContent = (appts: AppointmentRecord[]): string => {
    if (appts.length === 0) return "<div>No appointment data</div>";

    const isSingleAppointment = appts.length === 1;

    return `
    <div style="
      padding: 16px;
      min-width: 280px;
      max-width: 420px;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background: rgba(32, 32, 32, 0.95);
      backdrop-filter: blur(8px);
      border-radius: 12px;
      box-shadow: 0 8px 20px rgba(0,0,0,0.4);
      color: #f1f1f1;
      overflow: hidden;
    ">
      ${
        !isSingleAppointment
          ? `<div style="
              background: rgba(0, 120, 212, 0.15);
              border-left: 4px solid #00d4ff;
              padding: 10px 12px;
              margin-bottom: 14px;
              border-radius: 6px;
              font-size: 12px;
              font-weight: 600;
              color: #00d4ff;
            ">
              📌 ${appts.length} appointments at this location
            </div>`
          : ""
      }

      <div style="max-height: 400px; overflow-y: auto;">
        ${appts
          .map((appt, index) => {
            const startTime = formatDateTime(appt.scheduledstart);
            const endTime = formatDateTime(appt.scheduledend);
            const duration = calculateDuration(
              appt.scheduledstart,
              appt.scheduledend
            );
            const description = appt.description || "No description available";
            const regarding = appt.regardingobjectidname || "";

            return `
              <div style="
                margin-bottom: 16px;
                ${
                  index > 0
                    ? "border-top: 1px solid rgba(255,255,255,0.1); padding-top: 14px;"
                    : ""
                }
              ">
                <h4 style="margin: 0 0 10px 0; color: #00d4ff; font-size: 16px; font-weight: 600;">
                  ${escapeHtml(appt.subject)}
                </h4>

                <div style="
                  background: rgba(255,255,255,0.05);
                  padding: 10px;
                  border-radius: 6px;
                  font-size: 13px;
                  margin-bottom: 10px;
                  color: #e0e0e0;
                ">
                  <div style="margin-bottom: 4px;">
                    <span style="font-weight: 600;">📅</span>
                    <span style="margin-left: 6px;">${escapeHtml(
                      startTime
                    )}</span>
                  </div>
                  <div style="margin-bottom: 4px;">
                    <span style="font-weight: 600;">🕐</span>
                    <span style="margin-left: 6px;">${escapeHtml(
                      endTime
                    )}</span>
                  </div>
                  <div>
                    <span style="font-weight: 600;">⏱️</span>
                    <span style="margin-left: 6px;">${duration}</span>
                  </div>
                </div>

                ${
                  regarding
                    ? `<div style="margin-bottom: 8px; font-size: 13px;">
                        <span style="font-weight: 600;">👤 Regarding:</span>
                        <div style="margin-top: 2px; padding-left: 20px; color: #ccc;">${escapeHtml(
                          regarding
                        )}</div>
                      </div>`
                    : ""
                }

                ${
                  description && description !== "No description available"
                    ? `<div style="
                        font-size: 12px;
                        color: #ccc;
                        padding: 8px;
                        background: rgba(255,255,255,0.05);
                        border-radius: 6px;
                        max-height: 80px;
                        overflow-y: auto;
                      ">
                        ${escapeHtml(description)}
                      </div>`
                    : ""
                }
              </div>
            `;
          })
          .join("")}
      </div>
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

    markersRef.current.forEach((markerInfo) => {
      map.markers.remove(markerInfo.marker);
    });
    markersRef.current = [];

    // Step 1: Collect all addresses in parallel
    const addressPromises = appointments.map(async (appt) => {
      if (!appt.regardingobjectid) return null;
      const address = await fetchRegardingAddress(appt.regardingobjectid);
      return { appt, address };
    });

    const addressResults = await Promise.all(addressPromises);

    if (abortControllerRef.current?.signal.aborted) return;

    // Step 2: Group appointments by address
    const addressMap = new Map<string, AppointmentRecord[]>();
    for (const result of addressResults) {
      if (result && result.address) {
        if (!addressMap.has(result.address)) {
          addressMap.set(result.address, []);
        }
        addressMap.get(result.address)!.push(result.appt);
      }
    }

    // Step 3: Batch geocode unique addresses
    const uniqueAddresses = Array.from(addressMap.keys());
    const geocodeResults = await batchGeocodeAddresses(uniqueAddresses);

    if (abortControllerRef.current?.signal.aborted) return;

    // Step 4: Create markers
    const positions: atlas.data.Position[] = [];
    let markerCount = 0;

    for (const [address, appts] of addressMap) {
      const position = geocodeResults.get(address);

      if (!position) continue;

      positions.push(position);
      markerCount++;

      const marker = new atlas.HtmlMarker({
        color: "DodgerBlue",
        text: markerCount.toString(),
        position: position,
      });

      map.events.add("click", marker, () => {
        popup.close();
        popup.setOptions({
          content: createPopupContent(appts),
          position: position,
        });
        popup.open(map);
      });

      map.markers.add(marker);
      markersRef.current.push({ marker, appointment: appts[0] });
    }

    // Add user location to positions array if available
    if (userLocation?.position) {
      positions.push(userLocation.position);
    }

    if (positions.length > 0) {
      setErrorMessage("");
      const bounds = atlas.data.BoundingBox.fromPositions(positions);
      map.setCamera({
        bounds: bounds,
        padding: 80,
        maxZoom: 15,
      });
    }
  };

  const handleDueFilterChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    onFilterChange({
      dueFilter: e.target.value as DueFilter,
      searchText: searchText,
    });
  };

  const handleSearchChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const newSearchText = e.target.value;
    setSearchText(newSearchText);

    const timeoutId = setTimeout(() => {
      onFilterChange({
        dueFilter: currentFilter.dueFilter,
        searchText: newSearchText,
      });
    }, 300);

    return () => clearTimeout(timeoutId);
  };

  const handleRefreshClick = () => {
    onRefresh();
  };

  return (
    <div
      style={{
        width: "100%",
        height: "100%",
        display: "flex",
        flexDirection: "column",
        minHeight: "500px",
      }}
    >
      {/* Filter Controls Bar */}
      <div
        style={{
          background: "rgba(255, 255, 255, 0.97)",
          padding: "12px 16px",
          boxShadow: "0 2px 8px rgba(0,0,0,0.15)",
          fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
          display: "flex",
          gap: "12px",
          alignItems: "center",
          flexWrap: "wrap",
          borderBottom: "1px solid #e0e0e0",
          flexShrink: 0,
        }}
      >
        <div style={{ flex: "0 0 auto" }}>
          <div style={{ fontSize: "11px", color: "#666", marginBottom: "4px" }}>
            Showing appointments for
          </div>
          <div
            style={{ fontSize: "14px", fontWeight: "600", color: "#0078d4" }}
          >
            👤 {currentUserName}
          </div>
        </div>

        <div style={{ flex: "0 0 auto", minWidth: "180px" }}>
          <label
            style={{
              fontSize: "11px",
              color: "#666",
              display: "block",
              marginBottom: "4px",
            }}
          >
            Due
          </label>
          <select
            value={currentFilter.dueFilter}
            onChange={handleDueFilterChange}
            style={{
              width: "100%",
              padding: "6px 8px",
              fontSize: "13px",
              border: "1px solid #ccc",
              borderRadius: "4px",
              backgroundColor: "white",
              cursor: "pointer",
            }}
          >
            <option value="all">All</option>
            <option value="overdue">Overdue</option>
            <option value="today">Today or earlier</option>
            <option value="tomorrow">Tomorrow or earlier</option>
            <option value="next7days">Next 7 days or earlier</option>
            <option value="next30days">Next 30 days or earlier</option>
            <option value="next90days">Next 90 days or earlier</option>
            <option value="next6months">Next 6 months or earlier</option>
            <option value="next12months">Next 12 months or earlier</option>
          </select>
        </div>

        <div style={{ flex: "1 1 auto", minWidth: "200px" }}>
          <label
            style={{
              fontSize: "11px",
              color: "#666",
              display: "block",
              marginBottom: "4px",
            }}
          >
            Search
          </label>
          <input
            type="text"
            placeholder="Search appointments..."
            value={searchText}
            onChange={handleSearchChange}
            style={{
              width: "100%",
              padding: "6px 8px",
              fontSize: "13px",
              border: "1px solid #ccc",
              borderRadius: "4px",
            }}
          />
        </div>

        <div
          style={{
            flex: "0 0 auto",
            display: "flex",
            alignItems: "flex-end",
            gap: "8px",
          }}
        >
          <button
            onClick={handleRefreshClick}
            style={{
              padding: "6px 14px",
              fontSize: "13px",
              background: "linear-gradient(90deg, #4facfe 0%, #00f2fe 100%)",
              color: "#fff",
              border: "none",
              borderRadius: "6px",
              cursor: "pointer",
              fontWeight: "600",
              boxShadow: "0 4px 10px rgba(0,0,0,0.12)",
              transition: "transform 0.2s ease, box-shadow 0.2s ease",
            }}
            onMouseOver={(e) =>
              (e.currentTarget.style.transform = "scale(1.05)")
            }
            onMouseOut={(e) => (e.currentTarget.style.transform = "scale(1)")}
          >
            🔄 Refresh
          </button>
          <div
            style={{ fontSize: "11px", color: "#888", paddingBottom: "6px" }}
          >
            {appointments.length} of {allAppointmentsCount}
          </div>
        </div>
      </div>

      {/* Map Container */}
      <div
        ref={mapRef}
        style={{
          width: "100%",
          flex: "1",
          position: "relative",
          overflow: "hidden",
        }}
      >
        {/* No Results Message */}
        {!isLoading && !errorMessage && appointments.length === 0 && (
          <div
            style={{
              position: "absolute",
              top: "50%",
              left: "50%",
              transform: "translate(-50%, -50%)",
              background: "#ffffff",
              padding: "25px 30px",
              borderRadius: "6px",
              boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
              textAlign: "center",
              maxWidth: "calc(100% - 60px)",
              width: "380px",
              zIndex: 999,
              border: "1px solid #e8e8e8",
              pointerEvents: "auto",
            }}
          >
            <div
              style={{
                fontSize: "36px",
                marginBottom: "12px",
                lineHeight: "1",
              }}
            >
              📅
            </div>
            <h3
              style={{
                color: "#1a1a1a",
                marginTop: 0,
                marginBottom: "8px",
                fontSize: "17px",
                fontWeight: "700",
              }}
            >
              No Appointments Found
            </h3>
            <p
              style={{
                color: "#666666",
                margin: "10px 0 0 0",
                fontSize: "13px",
                lineHeight: "1.5",
              }}
            >
              {allAppointmentsCount === 0
                ? "No appointments available."
                : "No appointments match the current filter criteria."}
            </p>
            {allAppointmentsCount > 0 && (
              <p
                style={{
                  color: "#999999",
                  margin: "8px 0 0 0",
                  fontSize: "12px",
                }}
              >
                ({allAppointmentsCount} total appointments available)
              </p>
            )}
          </div>
        )}

        {/* Loading Indicator */}
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
            <p style={{ margin: 0, color: "#333", fontSize: "14px" }}>
              Loading your appointments...
            </p>
          </div>
        )}

        {/* Error Message */}
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
            <h3
              style={{
                color: "#d13438",
                marginTop: 0,
                marginBottom: "10px",
                fontSize: "18px",
              }}
            >
              ⚠️ Configuration Required
            </h3>
            <p
              style={{
                color: "#555",
                margin: 0,
                fontSize: "14px",
                lineHeight: "1.5",
              }}
            >
              {errorMessage}
            </p>
            <p
              style={{ color: "#777", margin: "15px 0 0 0", fontSize: "12px" }}
            >
              Please configure your Azure Maps subscription key in the control
              properties.
            </p>
          </div>
        )}
      </div>

      <style>{`
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      `}</style>
    </div>
  );
};

export default MapComponent;
