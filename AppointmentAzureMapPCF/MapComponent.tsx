import * as React from "react";
import * as atlas from "azure-maps-control";
import "azure-maps-control/dist/atlas.min.css";
import "./css/AppointmentAzureMapPCF.css";
import { IInputs } from "./generated/ManifestTypes";
import {
  AppointmentRecord,
  DueFilter,
  FilterOptions,
  UserLocation,
  DateRange,
} from "./types";
import { AzureMapsRouteService, RoutePoint, RouteResult } from "./RouteService";
import DatePicker from "./components/Datepicker";

interface MapComponentProps {
  appointments: AppointmentRecord[];
  allAppointmentsCount: number;
  azureMapsKey: string;
  context: ComponentFramework.Context<IInputs>;
  currentUserName: string;
  currentUserAddress: string;
  currentFilter: FilterOptions;
  showRoute: boolean;
  onFilterChange: (filter: FilterOptions) => void;
  onRouteToggle: (enabled: boolean) => void;
  onRefresh: () => void;
}

type GeocodeCache = Record<string, atlas.data.Position | null>;

interface MarkerInfo {
  marker: atlas.HtmlMarker;
  appointment: AppointmentRecord;
  position: atlas.data.Position;
  address: string;
}

interface CachedRoute {
  result: RouteResult;
  timestamp: number;
}

const AppointmentStatus = {
  Open: 0,
  Completed: 1,
  Canceled: 2,
  Scheduled: 3,
} as const;

const MapComponent: React.FC<MapComponentProps> = ({
  appointments,
  allAppointmentsCount,
  azureMapsKey,
  context,
  currentUserName,
  currentUserAddress,
  currentFilter,
  showRoute,
  onFilterChange,
  onRouteToggle,
  onRefresh,
}) => {
  const mapRef = React.useRef<HTMLDivElement | null>(null);
  const mapInstanceRef = React.useRef<atlas.Map | null>(null);
  const popupRef = React.useRef<atlas.Popup | null>(null);
  const geocodeCacheRef = React.useRef<GeocodeCache>({});
  const markersRef = React.useRef<MarkerInfo[]>([]);
  const userMarkerRef = React.useRef<atlas.HtmlMarker | null>(null);
  const abortControllerRef = React.useRef<AbortController | null>(null);
  const routeLayerRef = React.useRef<atlas.layer.LineLayer | null>(null);
  const routeSourceRef = React.useRef<atlas.source.DataSource | null>(null);
  const routeServiceRef = React.useRef<AzureMapsRouteService | null>(null);
  const routeCacheRef = React.useRef<Map<string, CachedRoute>>(new Map());

  const ROUTE_CACHE_DURATION = 5 * 60 * 1000; // 5 minutes
  const DISTANCE_THRESHOLD = 0.0001; // ~10 meters in degrees

  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [errorMessage, setErrorMessage] = React.useState<string>("");
  const [userLocation, setUserLocation] = React.useState<UserLocation | null>(
    null
  );
  const [routeData, setRouteData] = React.useState<RouteResult | null>(null);
  const [isCalculatingRoute, setIsCalculatingRoute] =
    React.useState<boolean>(false);
  const [showDatePicker, setShowDatePicker] = React.useState<boolean>(false);
  const [updatingAppointments, setUpdatingAppointments] = React.useState<
    Set<string>
  >(new Set());
  const [showCloseDialog, setShowCloseDialog] = React.useState<boolean>(false);
  const [selectedAppointment, setSelectedAppointment] =
    React.useState<AppointmentRecord | null>(null);
  const [selectedStatus, setSelectedStatus] = React.useState<number>(
    AppointmentStatus.Completed
  );

  const calculateDistance = (
    pos1: atlas.data.Position,
    pos2: atlas.data.Position
  ): number => {
    const dx = pos1[0] - pos2[0];
    const dy = pos1[1] - pos2[1];
    return Math.sqrt(dx * dx + dy * dy);
  };

  const positionsAreEqual = (
    pos1: atlas.data.Position,
    pos2: atlas.data.Position
  ): boolean => {
    return calculateDistance(pos1, pos2) < DISTANCE_THRESHOLD;
  };

  const updateAppointmentStatus = React.useCallback(
    async (appointmentId: string, newStateCode: number) => {
      if (
        newStateCode !== AppointmentStatus.Completed &&
        newStateCode !== AppointmentStatus.Canceled
      ) {
        console.warn(
          `[Status Update] Invalid state transition attempted: ${newStateCode}`
        );
        return;
      }

      setUpdatingAppointments((prev) => new Set(prev).add(appointmentId));
      setShowCloseDialog(false);

      try {
        const updateData: Record<string, number> = {
          statecode: newStateCode,
        };

        await context.webAPI.updateRecord(
          "appointment",
          appointmentId,
          updateData
        );

        popupRef.current?.close();

        setTimeout(() => {
          onRefresh();
        }, 500);
      } catch (error) {
        console.error(
          `[Status Update] ✗ Failed to close appointment ${appointmentId}:`,
          error
        );
        alert("Failed to close appointment. Please try again.");
      } finally {
        setUpdatingAppointments((prev) => {
          const newSet = new Set(prev);
          newSet.delete(appointmentId);
          return newSet;
        });
        setSelectedAppointment(null);
      }
    },
    [context, onRefresh]
  );

  const handleCloseAppointment = React.useCallback(
    (appointment: AppointmentRecord) => {
      setSelectedAppointment(appointment);
      setSelectedStatus(AppointmentStatus.Completed);
      setShowCloseDialog(true);
    },
    []
  );

  const handleCloseDialogConfirm = React.useCallback(() => {
    if (selectedAppointment) {
      void updateAppointmentStatus(selectedAppointment.id, selectedStatus);
    }
  }, [selectedAppointment, selectedStatus, updateAppointmentStatus]);

  const handleCloseDialogCancel = React.useCallback(() => {
    setShowCloseDialog(false);
    setSelectedAppointment(null);
    setSelectedStatus(AppointmentStatus.Completed);
  }, []);

  const openBingMapDirections = React.useCallback(
    (destinationAddress: string, appointmentSubject: string) => {
      if (!destinationAddress || destinationAddress.trim() === "") {
        console.warn("[Directions] No destination address available");
        alert("No address available");
        return;
      }

      try {
        const encodedAddress = encodeURIComponent(destinationAddress.trim());
        const bingMapsUrl = `https://www.bing.com/maps?q=${encodedAddress}`;

        const newWindow = window.open(bingMapsUrl, "_blank");

        if (newWindow === null) {
          console.error("[Directions] Pop-up was blocked by browser");
          alert("Pop-up blocked! Please allow pop-ups for this site.");
        }
      } catch (error) {
        console.error("[Directions] Error opening Bing Maps:", error);
        alert("Error opening Bing Maps. Check console for details.");
      }
    },
    []
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

    if (!currentUserAddress) {
      setErrorMessage("User address is required to display the map");
      setIsLoading(false);
      return;
    }

    setUserLocation(null);
    initializeMap();

    return () => {
      cleanup();
    };
  }, [azureMapsKey, currentUserAddress]);

  React.useEffect(() => {
    if (
      mapInstanceRef.current &&
      popupRef.current &&
      !isLoading &&
      userLocation
    ) {
      abortControllerRef.current = new AbortController();
      void updateMarkers();
    }
    return () => {
      abortControllerRef.current?.abort();
    };
  }, [appointments, showRoute, userLocation, isLoading]);

  React.useEffect(() => {
    const handlePopupClick = (e: MouseEvent) => {
      const target = e.target as HTMLElement;

      if (target.classList.contains("appointment-popup-directions-btn")) {
        e.preventDefault();
        e.stopPropagation();

        const address = target.getAttribute("data-address");
        const subject = target.getAttribute("data-subject");

        if (address && subject) {
          openBingMapDirections(address, subject);
        }
      }

      if (target.classList.contains("appointment-popup-close-btn")) {
        e.preventDefault();
        e.stopPropagation();

        const appointmentId = target.getAttribute("data-appointment-id");

        if (appointmentId) {
          const appointment = appointments.find((a) => a.id === appointmentId);
          if (appointment) {
            handleCloseAppointment(appointment);
          }
        }
      }
    };

    document.addEventListener("click", handlePopupClick, true);

    return () => {
      document.removeEventListener("click", handlePopupClick, true);
    };
  }, [openBingMapDirections, handleCloseAppointment, appointments]);

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

    if (markersRef.current.length > 0 && mapInstanceRef.current) {
      markersRef.current.forEach((m) => {
        mapInstanceRef.current?.markers.remove(m.marker);
      });
      markersRef.current = [];
    }

    if (routeSourceRef.current) {
      routeSourceRef.current.clear();
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
        minZoom: 2,
        maxZoom: 20,
      });

      mapInstanceRef.current = map;
      routeServiceRef.current = new AzureMapsRouteService(azureMapsKey);

      const popup = new atlas.Popup({
        pixelOffset: [0, -18],
        closeButton: true,
      });
      popupRef.current = popup;

      map.events.add("ready", async () => {
        const routeSource = new atlas.source.DataSource();
        map.sources.add(routeSource);
        routeSourceRef.current = routeSource;

        const routeLayer = new atlas.layer.LineLayer(routeSource, undefined, {
          strokeColor: "rgba(79, 172, 254, 0.8)",
          strokeWidth: 3,
          lineJoin: "round",
          lineCap: "round",
        });
        map.layers.add(routeLayer);
        routeLayerRef.current = routeLayer;

        map.controls.add(
          [
            new atlas.control.ZoomControl(),
            new atlas.control.CompassControl(),
            new atlas.control.PitchControl(),
          ],
          {
            position: atlas.ControlPosition.TopRight,
          }
        );

        map.setCamera({
          bounds: [-180, -85, 180, 85],
          padding: 0,
        });

        if (currentUserAddress) {
          await geocodeAndDisplayUserLocation();
        }

        setIsLoading(false);
      });

      map.events.add("error", (error) => {
        console.error("[Map] Error event:", error);
        if (markersRef.current.length === 0) {
          setErrorMessage("Failed to load map");
        }
        setIsLoading(false);
      });
    } catch (error) {
      console.error("[Map] Initialization error:", error);
      setErrorMessage("Failed to initialize map");
      setIsLoading(false);
    }
  };

  const geocodeAddress = async (
    address: string
  ): Promise<atlas.data.Position | null> => {
    if (!address || !address.trim()) return null;
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
          const position: atlas.data.Position = [
            result.position.lon,
            result.position.lat,
          ];
          geocodeCacheRef.current[normalizedAddress] = position;
          return position;
        }
      }

      geocodeCacheRef.current[normalizedAddress] = null;
      console.warn(`[Geocode] No results for: ${address}`);
      return null;
    } catch (error) {
      if ((error as Error).name !== "AbortError") {
        console.error("[Geocode] Error geocoding address:", address, error);
        geocodeCacheRef.current[normalizedAddress] = null;
      }
      return null;
    }
  };

  const batchGeocodeAddresses = async (
    addresses: string[]
  ): Promise<Map<string, atlas.data.Position | null>> => {
    const results = new Map<string, atlas.data.Position | null>();
    const batchSize = 10;
    const delayMs = 50;

    for (let i = 0; i < addresses.length; i += batchSize) {
      if (abortControllerRef.current?.signal.aborted) break;
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

  const geocodeAndDisplayUserLocation = async () => {
    const map = mapInstanceRef.current;
    const popup = popupRef.current;

    if (!map || !popup || !currentUserAddress) return;

    if (userMarkerRef.current) {
      map.markers.remove(userMarkerRef.current);
      userMarkerRef.current = null;
    }

    const position = await geocodeAddress(currentUserAddress);

    if (position) {
      const newUserLocation = { address: currentUserAddress, position };
      setUserLocation(newUserLocation);

      const userMarker = new atlas.HtmlMarker({
        color: "red",
        position,
      });

      map.events.add("click", userMarker, () => {
        popup.close();
        popup.setOptions({
          content: createUserPopupContent(currentUserName, currentUserAddress),
          position,
        });
        popup.open(map);
      });

      map.markers.add(userMarker);
      userMarkerRef.current = userMarker;
    } else {
      console.warn("[User Location] Failed to geocode user address");
      setUserLocation(null);
      setErrorMessage(
        `Unable to locate your address: "${currentUserAddress}". Please update your address in the system user settings.`
      );
      setIsLoading(false);
    }
  };

  const fetchRegardingAddress = async (
    regardingobjectid: ComponentFramework.EntityReference
  ): Promise<string | null> => {
    if (!regardingobjectid?.id || !regardingobjectid?.etn) return null;

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
      console.error("[Fetch Address] Error fetching regarding address:", error);
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

  const escapeHtml = (text: string): string => {
    if (!text) return "";
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
  };

  const formatRouteDuration = (durationInSeconds: number): string => {
    const totalMinutes = Math.round(durationInSeconds / 60);
    if (totalMinutes >= 60) {
      const hours = Math.floor(totalMinutes / 60);
      const minutes = totalMinutes % 60;
      return minutes > 0 ? `${hours}h ${minutes}m` : `${hours}h`;
    }
    return `${totalMinutes}m`;
  };

  const formatSelectedDateRange = (): string => {
    if (currentFilter.customDateRange) {
      const start = currentFilter.customDateRange.startDate.toLocaleDateString(
        "en-US",
        {
          month: "short",
          day: "numeric",
          year: "numeric",
        }
      );
      const end = currentFilter.customDateRange.endDate.toLocaleDateString(
        "en-US",
        {
          month: "short",
          day: "numeric",
          year: "numeric",
        }
      );
      return `${start} - ${end}`;
    }
    return "";
  };

  const createUserPopupContent = (
    userName: string,
    address: string
  ): string => {
    return `
      <div class="user-popup-container">
        <div class="user-popup-header">
          <svg xmlns="http://www.w3.org/2000/svg" class="user-popup-icon" fill="#ff0000" viewBox="0 0 24 24">
            <path d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5c-1.38 0-2.5-1.12-2.5-2.5S10.62 6.5 12 6.5s2.5 1.12 2.5 2.5S13.38 11.5 12 11.5z"/>
          </svg>
          <h3 class="user-popup-title">Your Location</h3>
        </div>
        <div class="user-popup-field">
          <div class="user-popup-field-label">User</div>
          <div class="user-popup-field-value">${escapeHtml(userName)}</div>
        </div>
        <div>
          <div class="user-popup-field-label">Address</div>
          <div class="user-popup-address">${escapeHtml(address)}</div>
        </div>
      </div>
    `;
  };

  const createPopupContent = (
    appts: AppointmentRecord[],
    address: string
  ): string => {
    if (appts.length === 0) return "<div>No appointment data</div>";
    const isSingleAppointment = appts.length === 1;

    return `
      <div class="appointment-popup-container">
        ${!isSingleAppointment
        ? `<div class="appointment-popup-badge">📌 ${appts.length} appointments at this location</div>`
        : ""
      }
        <div class="appointment-popup-list">
          ${appts
        .map((appt, index) => {
          const startTime = formatDateTime(appt.scheduledstart);
          const endTime = formatDateTime(appt.scheduledend);
          const duration = calculateDuration(
            appt.scheduledstart,
            appt.scheduledend
          );
          const regarding = appt.regardingobjectidname || "";
          const isUpdating = updatingAppointments.has(appt.id);

          return `
                <div class="appointment-popup-item ${index > 0 ? "appointment-popup-item-separator" : ""
            }">
                  <h4 class="appointment-popup-subject">${escapeHtml(
              appt.subject
            )}</h4>
                  <div class="appointment-popup-details">
                    <div class="appointment-popup-detail-row">
                      <span class="appointment-popup-detail-icon">📅</span>
                      <span class="appointment-popup-detail-text">${escapeHtml(
              startTime
            )}</span>
                    </div>
                    <div class="appointment-popup-detail-row">
                      <span class="appointment-popup-detail-icon">🕐</span>
                      <span class="appointment-popup-detail-text">${escapeHtml(
              endTime
            )}</span>
                    </div>
                    <div class="appointment-popup-detail-row">
                      <span class="appointment-popup-detail-icon">⏱️</span>
                      <span class="appointment-popup-detail-text">${duration}</span>
                    </div>
                  </div>
                  ${regarding
              ? `<div class="appointment-popup-regarding"><span class="appointment-popup-regarding-label">👤 Regarding:</span><div class="appointment-popup-regarding-value">${escapeHtml(
                regarding
              )}</div></div>`
              : ""
            }
                  <div class="appointment-popup-actions">
                    <button 
                      class="appointment-popup-directions-btn" 
                      data-address="${escapeHtml(address)}" 
                      data-subject="${escapeHtml(appt.subject)}"
                      ${isUpdating ? "disabled" : ""}
                    >
                      Get Directions
                    </button>
                    <button 
                      class="appointment-popup-close-btn" 
                      data-appointment-id="${escapeHtml(appt.id)}" 
                      data-appointment-subject="${escapeHtml(appt.subject)}"
                      ${isUpdating ? "disabled" : ""}
                    >
                      ${isUpdating ? "⏳ Updating..." : "✕ Close Meeting"}
                    </button>
                  </div>
                </div>
              `;
        })
        .join("")}
        </div>
      </div>
    `;
  };

  const fitMapToMarkers = () => {
    const map = mapInstanceRef.current;
    if (!map) return;

    const positions: atlas.data.Position[] = [];

    if (userLocation?.position) {
      positions.push(userLocation.position);
    }

    markersRef.current.forEach((markerInfo) => {
      positions.push(markerInfo.position);
    });

    if (positions.length > 0) {
      const bounds = atlas.data.BoundingBox.fromPositions(positions);
      map.setCamera({
        bounds,
        padding: 80,
        maxZoom: 15,
        minZoom: 1,
      });
    }
  };

  const updateMarkers = async () => {
    const map = mapInstanceRef.current;
    const popup = popupRef.current;
    if (!map || !popup) return;

    if (!userLocation?.position) {
      return;
    }

    popup.close();

    markersRef.current.forEach((m) => {
      map.markers.remove(m.marker);
    });
    markersRef.current = [];

    if (routeSourceRef.current) routeSourceRef.current.clear();
    setRouteData(null);

    if (appointments.length === 0) {
      fitMapToMarkers();
      return;
    }

    const addressPromises = appointments.map(async (appt) => {
      let address: string | null = null;

      if (appt.location && appt.location.trim()) {
        const locationAddress = appt.location.trim();
        return { appt, address: locationAddress, source: "location" };
      }

      if (appt.regardingobjectid) {
        address = await fetchRegardingAddress(appt.regardingobjectid);
        if (address && address.trim()) {
          return { appt, address: address.trim(), source: "regarding" };
        }
      }

      return null;
    });

    const addressResults = await Promise.all(addressPromises);
    if (abortControllerRef.current?.signal.aborted) return;

    const addressMap = new Map<string, AppointmentRecord[]>();
    for (const r of addressResults) {
      if (r && r.address && r.address.trim()) {
        const normalizedAddress = r.address.trim();
        if (!addressMap.has(normalizedAddress))
          addressMap.set(normalizedAddress, []);
        addressMap.get(normalizedAddress)!.push(r.appt);
      }
    }

    const uniqueAddresses = Array.from(addressMap.keys());
    if (uniqueAddresses.length === 0) {
      fitMapToMarkers();
      return;
    }

    const geocodeResults = await Promise.all(
      uniqueAddresses.map((addr) =>
        geocodeAddress(addr).then((pos) => ({ address: addr, position: pos }))
      )
    );

    if (abortControllerRef.current?.signal.aborted) return;

    const geocodeMap = new Map(
      geocodeResults
        .filter((r) => r.position)
        .map((r) => [r.address, r.position])
    );

    const sortedAddresses = uniqueAddresses.sort((a, b) => {
      const apptA = addressMap.get(a)?.[0];
      const apptB = addressMap.get(b)?.[0];
      if (!apptA || !apptB) return 0;
      return apptA.scheduledstart.getTime() - apptB.scheduledstart.getTime();
    });

    const routePoints: RoutePoint[] = [];
    let markerCount = 0;

    for (const address of sortedAddresses) {
      const appts = addressMap.get(address);
      const position = geocodeMap.get(address);
      if (!position || !appts) continue;

      markerCount++;

      const isFirstMarker = markerCount === 1;

      const marker = new atlas.HtmlMarker({
        color: isFirstMarker ? "#E53935" : "DodgerBlue",
        text: markerCount.toString(),
        position: position,
      });

      map.events.add("click", marker, () => {
        popup.close();
        popup.setOptions({
          content: createPopupContent(appts, address),
          position: position,
        });
        popup.open(map);
      });

      map.markers.add(marker);
      markersRef.current.push({
        marker,
        appointment: appts[0],
        position,
        address,
      });

      routePoints.push({
        position,
        address,
        appointmentId: appts[0].id,
        subject: appts[0].subject,
        scheduledstart: appts[0].scheduledstart,
      });
    }

    if (showRoute && routePoints.length > 0 && userLocation?.position) {
      void calculateAndDisplayRoute(userLocation.position, routePoints);
    } else if (!showRoute && routeSourceRef.current) {
      routeSourceRef.current.clear();
      setRouteData(null);
    }

    fitMapToMarkers();
  };

  const getCacheKey = (startPos: atlas.data.Position, points: RoutePoint[]) =>
    JSON.stringify([startPos, points.map((p) => p.position)]);

  const displayRouteOnMap = (result: RouteResult) => {
    if (!routeSourceRef.current) return;
    const routeFeature = new atlas.data.Feature(
      new atlas.data.LineString(result.routeCoordinates),
      {
        distance: result.totalDistance,
        duration: result.totalDuration,
      }
    );
    routeSourceRef.current.clear();
    routeSourceRef.current.add(routeFeature);
  };

  const calculateAndDisplayRoute = async (
    startPosition: atlas.data.Position,
    routePoints: RoutePoint[]
  ) => {
    if (!routeServiceRef.current || !routeSourceRef.current) return;

    const filteredPoints = routePoints.filter(
      (point) => !positionsAreEqual(point.position, startPosition)
    );

    if (filteredPoints.length === 0) {
      console.warn(
        "[Route] All appointments are at user location, skipping route calculation"
      );
      setRouteData(null);
      setIsCalculatingRoute(false);
      return;
    }

    setIsCalculatingRoute(true);
    routeSourceRef.current.clear();
    setRouteData(null);

    try {
      const cacheKey = getCacheKey(startPosition, filteredPoints);
      const cached = routeCacheRef.current.get(cacheKey);
      if (cached && Date.now() - cached.timestamp < ROUTE_CACHE_DURATION) {
        setRouteData(cached.result);
        displayRouteOnMap(cached.result);
        setIsCalculatingRoute(false);
        return;
      }

      const result = await routeServiceRef.current.calculateChronologicalRoute(
        startPosition,
        filteredPoints
      );

      if (
        result &&
        result.routeCoordinates &&
        result.routeCoordinates.length > 0
      ) {
        routeCacheRef.current.set(cacheKey, { result, timestamp: Date.now() });
        setRouteData(result);
        displayRouteOnMap(result);
      } else {
        routeSourceRef.current.clear();
        setRouteData(null);
        console.warn("[Route] No route data returned from API");
      }
    } catch (error) {
      console.error("[Route] Failed to calculate route:", error);
      routeSourceRef.current.clear();
      setRouteData(null);
    } finally {
      setIsCalculatingRoute(false);
    }
  };

  const handleDueFilterChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const newValue = e.target.value as DueFilter;

    if (newValue === "customDateRange") {
      setShowDatePicker(true);
    } else {
      onFilterChange({ dueFilter: newValue, customDateRange: undefined });
    }
  };

  const handleDateRangeSelect = (range: DateRange) => {
    onFilterChange({
      dueFilter: "customDateRange",
      customDateRange: range,
    });
  };

  const handleDatePickerClose = () => {
    setShowDatePicker(false);
  };

  const handleRefreshClick = () => {
    onRefresh();
  };

  const handleRouteToggleClick = () => {
    const newShowRoute = !showRoute;
    onRouteToggle(newShowRoute);
  };

  const handleDateBadgeClick = React.useCallback((e: React.MouseEvent) => {
    e.stopPropagation();
    
    const target = e.target as HTMLElement;
    if (!target.classList.contains('clear-date-btn') && 
        !target.closest('.clear-date-btn')) {
      setShowDatePicker(true);
    }
  }, []);

  return (
    <div className="appointment-azure-map-container">
      <div className="filter-controls-wrapper">
        <div className="filter-controls-bar">
          <div className="user-info-section">
            <div className="user-info-label">Showing appointments for</div>
            <div className="user-info-name">👤 {currentUserName}</div>
          </div>

          <div className="filter-section">
            <label className="filter-label">Due</label>
            <select
              value={currentFilter.dueFilter}
              onChange={handleDueFilterChange}
              className="filter-select"
            >
              <option value="today">Today</option>
              <option value="tomorrow">Tomorrow</option>
              <option value="next7days">Next 7 days</option>
              <option value="next30days">Next 30 days</option>
              <option value="next90days">Next 90 days</option>
              <option value="next6months">Next 6 months</option>
              <option value="next12months">Next 12 months</option>
              <option value="overdue">Overdue</option>
              <option value="customDateRange">Custom Date Range</option>
              <option value="all">All</option>
            </select>

            {currentFilter.dueFilter === "customDateRange" &&
              currentFilter.customDateRange && (
                <div 
                  className="selected-date-badge"
                  onClick={handleDateBadgeClick}
                  role="button"
                  tabIndex={0}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault();
                      setShowDatePicker(true);
                    }
                  }}
                  title="Click to edit date range"
                  aria-label="Edit date range"
                >
                  <span className="selected-date-text">
                    {formatSelectedDateRange()}
                  </span>
                  <button
                    className="clear-date-btn"
                    onClick={(e) => {
                      e.stopPropagation();
                      onFilterChange({
                        dueFilter: "today",
                        customDateRange: undefined,
                      });
                    }}
                    title="Clear date range"
                    aria-label="Clear date range"
                  >
                    ✕
                  </button>
                </div>
              )}
          </div>

          <div className="route-toggle-section">
            {appointments.length > 0 && (
              <label className="route-toggle-container">
                <input
                  type="checkbox"
                  checked={showRoute}
                  onChange={handleRouteToggleClick}
                  className="route-toggle-checkbox"
                  disabled={isCalculatingRoute}
                />
                <span className="route-toggle-label">Show Route</span>
              </label>
            )}

            <div className="route-status-placeholder">
              {isCalculatingRoute ? (
                <div className="route-calculating-indicator">
                  <div className="route-calculating-spinner" />
                  <span className="route-calculating-text">
                    Optimizing route...
                  </span>
                </div>
              ) : routeData && showRoute ? (
                <div className="route-info">
                  <span className="route-info-item">
                    📏 {(routeData.totalDistance * 0.000621371).toFixed(1)} mi
                  </span>
                  <span className="route-info-item">
                    ⏱️ {formatRouteDuration(routeData.totalDuration)}
                  </span>
                </div>
              ) : null}
            </div>
          </div>

          <div className="action-section">
            <button onClick={handleRefreshClick} className="refresh-button">
              Refresh
            </button>
            <div className="appointment-count">
              {appointments.length} of {allAppointmentsCount}
            </div>
          </div>
        </div>
      </div>

      <div className="map-content-wrapper">
        <div ref={mapRef} className="map-container">
          {!isLoading && !errorMessage && appointments.length === 0 && (
            <div className="overlay-message no-appointments-overlay">
              <div className="no-appointments-icon">📅</div>
              <h3 className="no-appointments-title">No Appointments Found</h3>
              <p className="no-appointments-message">
                {allAppointmentsCount === 0
                  ? "No appointments available."
                  : currentFilter.dueFilter === "customDateRange" &&
                    currentFilter.customDateRange
                    ? `No appointments found for ${formatSelectedDateRange()}.`
                    : "No appointments match the current filter criteria."}
              </p>
            </div>
          )}

          {isLoading && (
            <div className="loading-overlay">
              <div className="loading-spinner" />
              <p className="loading-text">Loading your appointments...</p>
            </div>
          )}

          {errorMessage && (
            <div className="error-overlay">
              <h3 className="error-title">⚠️ Configuration Required</h3>
              <p className="error-message">{errorMessage}</p>
            </div>
          )}
        </div>
      </div>

      {showDatePicker && (
        <DatePicker
          dateRange={currentFilter.customDateRange}
          onDateRangeSelect={handleDateRangeSelect}
          onClose={handleDatePickerClose}
        />
      )}

      {showCloseDialog && selectedAppointment && (
        <div className="close-dialog-overlay" onClick={handleCloseDialogCancel}>
          <div
            className="close-dialog-container"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="close-dialog-header">
              <h3 className="close-dialog-title">Close Meeting</h3>
              <button
                className="close-dialog-close-btn"
                onClick={handleCloseDialogCancel}
                aria-label="Close dialog"
              >
                ✕
              </button>
            </div>

            <div className="close-dialog-content">
              <p className="close-dialog-question">
                Do you want to close the selected 1 Meeting?
              </p>
              <p className="close-dialog-instruction">
                Select the status of the closing Meeting.
              </p>

              <div className="close-dialog-field">
                <label className="close-dialog-label">State</label>
                <select
                  className="close-dialog-select"
                  value={selectedStatus}
                  onChange={(e) =>
                    setSelectedStatus(parseInt(e.target.value, 10))
                  }
                >
                  <option value={AppointmentStatus.Completed}>Completed</option>
                  <option value={AppointmentStatus.Canceled}>Canceled</option>
                </select>
              </div>
            </div>

            <div className="close-dialog-footer">
              <button
                className="close-dialog-btn close-dialog-confirm-btn"
                onClick={handleCloseDialogConfirm}
                disabled={updatingAppointments.has(selectedAppointment.id)}
              >
                {updatingAppointments.has(selectedAppointment.id)
                  ? "Updating..."
                  : "Close Meeting"}
              </button>
              <button
                className="close-dialog-btn close-dialog-cancel-btn"
                onClick={handleCloseDialogCancel}
                disabled={updatingAppointments.has(selectedAppointment.id)}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default MapComponent;