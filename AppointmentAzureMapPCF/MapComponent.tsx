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
import { AzureMapsRouteService, RoutePoint, RouteResult, RouteLeg } from "./RouteService";
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

interface AppointmentProperties {
  appointments: AppointmentRecord[];
  address: string;
  markerNumber: number;
  isFirstMarker: boolean;
  routeLegDistance?: number;
  routeLegDuration?: number;
  routeLegFrom?: string;
  routeLegTo?: string;
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
  const userMarkerRef = React.useRef<atlas.HtmlMarker | null>(null);
  const abortControllerRef = React.useRef<AbortController | null>(null);
  const routeLayerRef = React.useRef<atlas.layer.LineLayer | null>(null);
  const routeSourceRef = React.useRef<atlas.source.DataSource | null>(null);
  const routeServiceRef = React.useRef<AzureMapsRouteService | null>(null);
  const routeCacheRef = React.useRef<Map<string, CachedRoute>>(new Map());
  const appointmentSourceRef = React.useRef<atlas.source.DataSource | null>(null);
  const symbolLayerRef = React.useRef<atlas.layer.SymbolLayer | null>(null);
  const clusterBubbleLayerRef = React.useRef<atlas.layer.BubbleLayer | null>(null);
  const clusterSymbolLayerRef = React.useRef<atlas.layer.SymbolLayer | null>(null);
  const showRouteRef = React.useRef<boolean>(showRoute);

  const ROUTE_CACHE_DURATION = 5 * 60 * 1000;
  const DISTANCE_THRESHOLD = 0.0001;

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

  React.useEffect(() => {
    showRouteRef.current = showRoute;
  }, [showRoute]);

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
        alert("No address available");
        return;
      }

      try {
        const encodedAddress = encodeURIComponent(destinationAddress.trim());
        const bingMapsUrl = `https://www.bing.com/maps?q=${encodedAddress}`;

        const newWindow = window.open(bingMapsUrl, "_blank");

        if (newWindow === null) {
          alert("Pop-up blocked! Please allow pop-ups for this site.");
        }
      } catch (error) {
        alert("Error opening Bing Maps.");
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

    if (appointmentSourceRef.current) {
      appointmentSourceRef.current.clear();
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
          strokeColor: "#2272B9",
          strokeWidth: 4,
          lineJoin: "round",
          lineCap: "round",
        });
        map.layers.add(routeLayer);
        routeLayerRef.current = routeLayer;

        const appointmentSource = new atlas.source.DataSource(undefined, {
          cluster: true,
          clusterRadius: 50,
          clusterMaxZoom: 14,
          clusterProperties: {
            hasFirstMarker: ["any", ["get", "isFirstMarker"]]
          }
        });
        map.sources.add(appointmentSource);
        appointmentSourceRef.current = appointmentSource;

        const clusterBubbleLayer = new atlas.layer.BubbleLayer(
          appointmentSource,
          undefined,
          {
            radius: [
              "step",
              ["get", "point_count"],
              20,
              5, 25,
              10, 30
            ],
            color: [
              "case",
              ["get", "hasFirstMarker"],
              "#E53935",
              [
                "step",
                ["get", "point_count"],
                "#4FC2F6",
                5, "#3FA9D9",
                10, "#2E8AB8"
              ]
            ],
            strokeWidth: 0,
            filter: ["has", "point_count"],
          }
        );
        map.layers.add(clusterBubbleLayer);
        clusterBubbleLayerRef.current = clusterBubbleLayer;

        const clusterSymbolLayer = new atlas.layer.SymbolLayer(
          appointmentSource,
          undefined,
          {
            iconOptions: {
              image: "none",
            },
            textOptions: {
              textField: ["get", "point_count_abbreviated"],
              offset: [0, 0],
              color: "#FFFFFF",
              size: 14,
              font: ["SegoeUi-Bold"],
            },
            filter: ["has", "point_count"],
          }
        );
        map.layers.add(clusterSymbolLayer);
        clusterSymbolLayerRef.current = clusterSymbolLayer;

        const symbolLayer = new atlas.layer.SymbolLayer(
          appointmentSource,
          undefined,
          {
            iconOptions: {
              image: [
                "case",
                ["get", "isFirstMarker"],
                "marker-red",
                "marker-blue"
              ],
              allowOverlap: true,
              anchor: "bottom",
              size: 1.0,
            },
            textOptions: {
              textField: ["to-string", ["get", "markerNumber"]],
              offset: [0, -1.6],
              color: "#FFFFFF",
              size: 12,
              font: ["SegoeUi-Bold"],
              allowOverlap: true,
            },
            filter: ["!", ["has", "point_count"]],
          }
        );
        map.layers.add(symbolLayer);
        symbolLayerRef.current = symbolLayer;

        map.events.add("click", (e) => {
          const clusterFeatures = map.layers.getRenderedShapes(e.position, [
            clusterBubbleLayer,
            clusterSymbolLayer
          ]);

          if (clusterFeatures.length > 0) {
            const feature = clusterFeatures[0];
            const properties = (feature as unknown as { properties?: Record<string, unknown> }).properties;

            if (properties?.point_count) {
              onClusterClick(e);
              return;
            }
          }

          const appointmentFeatures = map.layers.getRenderedShapes(e.position, [
            symbolLayer
          ]);

          if (appointmentFeatures.length > 0) {
            const feature = appointmentFeatures[0];
            const shapeData = (feature as unknown as { data?: unknown }).data;
            const geoJsonFeature = shapeData as { properties?: AppointmentProperties };
            const properties = geoJsonFeature.properties;

            if (properties?.appointments) {
              onAppointmentClick(e);
            }
          }
        });

        map.events.add("mousemove", (e) => {
          const features = map.layers.getRenderedShapes(e.position, [
            clusterBubbleLayer,
            clusterSymbolLayer,
            symbolLayer
          ]);

          if (features.length > 0) {
            map.getCanvasContainer().style.cursor = "pointer";
          } else {
            map.getCanvasContainer().style.cursor = "grab";
          }
        });

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
        if (appointmentSourceRef.current?.getShapes().length === 0) {
          setErrorMessage("Failed to load map");
        }
        setIsLoading(false);
      });
    } catch (error) {
      setErrorMessage("Failed to initialize map");
      setIsLoading(false);
    }
  };

  const onClusterClick = (e: atlas.MapMouseEvent) => {
    if (!appointmentSourceRef.current || !mapInstanceRef.current) {
      return;
    }

    const features = mapInstanceRef.current.layers.getRenderedShapes(e.position, [
      clusterBubbleLayerRef.current!,
      clusterSymbolLayerRef.current!
    ]);

    if (features.length === 0) {
      return;
    }

    const feature = features[0];
    const geoJsonFeature = feature as unknown as {
      properties?: Record<string, unknown>,
      geometry?: { coordinates?: atlas.data.Position }
    };
    const properties = geoJsonFeature.properties;

    if (!properties?.point_count) {
      return;
    }

    const clusterId = properties.cluster_id;
    const clusterPosition: atlas.data.Position = geoJsonFeature.geometry?.coordinates as atlas.data.Position;

    if (clusterId && appointmentSourceRef.current) {
      appointmentSourceRef.current
        .getClusterExpansionZoom(clusterId as number)
        .then((zoom) => {
          const currentZoom = mapInstanceRef.current?.getCamera().zoom || 4;
          const targetZoom = Math.max(zoom, currentZoom + 1);

          mapInstanceRef.current?.setCamera({
            center: clusterPosition,
            zoom: targetZoom,
            type: "ease",
            duration: 300,
          });

          return undefined;
        })
        .catch(() => {
          const currentZoom = mapInstanceRef.current?.getCamera().zoom || 4;
          const targetZoom = currentZoom + 2;
          mapInstanceRef.current?.setCamera({
            center: clusterPosition,
            zoom: targetZoom,
            type: "ease",
            duration: 300,
          });
        });
    } else {
      const currentZoom = mapInstanceRef.current?.getCamera().zoom || 4;
      const targetZoom = currentZoom + 2;
      mapInstanceRef.current?.setCamera({
        center: clusterPosition,
        zoom: targetZoom,
        type: "ease",
        duration: 300,
      });
    }
  };

  const onAppointmentClick = (e: atlas.MapMouseEvent) => {
    if (!popupRef.current || !mapInstanceRef.current || !symbolLayerRef.current) return;

    const features = mapInstanceRef.current.layers.getRenderedShapes(e.position, [
      symbolLayerRef.current
    ]);

    if (features.length === 0) return;

    const feature = features[0];
    const shapeData = (feature as unknown as { data?: unknown }).data;
    const geoJsonFeature = shapeData as { properties?: AppointmentProperties, geometry?: { coordinates?: atlas.data.Position } };
    const properties = geoJsonFeature.properties;
    const position: atlas.data.Position = geoJsonFeature.geometry?.coordinates as atlas.data.Position;

    if (!properties) return;

    popupRef.current.close();
    popupRef.current.setOptions({
      content: createPopupContent(
        properties.appointments,
        properties.address,
        properties.routeLegDistance,
        properties.routeLegDuration,
        properties.routeLegFrom,
        properties.markerNumber,
        showRouteRef.current
      ),
      position: position,
    });
    popupRef.current.open(mapInstanceRef.current);
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

    if (userMarkerRef.current) {
      map.markers.remove(userMarkerRef.current);
      userMarkerRef.current = null;
    }

    const position = await geocodeAddress(currentUserAddress);

    if (position) {
      const newUserLocation = { address: currentUserAddress, position };
      setUserLocation(newUserLocation);

      const userMarker = new atlas.HtmlMarker({
        position,
        htmlContent: `
    <div class="user-location-ring"></div>
  `,
        anchor: "center",
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
    address: string,
    routeLegDistance?: number,
    routeLegDuration?: number,
    routeLegFrom?: string,
    stopNumber?: number,
    isRouteVisible?: boolean
  ): string => {
    if (appts.length === 0) return "<div>No appointment data</div>";

    const isSingleAppointment = appts.length === 1;
    const routeInfoHtml = routeLegDistance != null && routeLegDuration != null && isRouteVisible ? `
      <div class="route-info-container">
        <div class="route-info-stats">
          <div class="route-info-stat">
            <span>${!isNaN(Number(routeLegDistance)) ? (Number(routeLegDistance) * 0.000621371).toFixed(1) : '?'} mi</span>
          </div>
          <div class="route-info-stat">
            <span>${!isNaN(Number(routeLegDuration)) ? formatRouteDuration(Number(routeLegDuration)) : '?'}</span>
          </div>
        </div>
        <div class="route-info-from">
          From: ${escapeHtml(routeLegFrom || 'Unknown')}
        </div>
      </div>
    ` : '';

    return `
    <div class="appointment-popup-container">
      ${!isSingleAppointment
        ? `<div class="appointment-popup-badge">📌 ${appts.length} appointments at this location</div>`
        : ""
      }

      ${routeInfoHtml}

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
                <h4 class="appointment-popup-subject">
                  ${escapeHtml(appt.subject)}
                </h4>

                <div class="appointment-popup-details">
                  <div class="appointment-popup-detail-row">
                    <span class="appointment-popup-detail-icon">📅</span>
                    <span class="appointment-popup-detail-text">
                      ${escapeHtml(startTime)}
                    </span>
                  </div>

                  <div class="appointment-popup-detail-row">
                    <span class="appointment-popup-detail-icon">🕐</span>
                    <span class="appointment-popup-detail-text">
                      ${escapeHtml(endTime)}
                    </span>
                  </div>

                  <div class="appointment-popup-detail-row">
                    <span class="appointment-popup-detail-icon">⏱️</span>
                    <span class="appointment-popup-detail-text">
                      ${duration}
                    </span>
                  </div>

                  ${regarding
              ? `
                        <div class="appointment-popup-detail-row">
                          <span class="appointment-popup-detail-icon">👤</span>
                          <span class="appointment-popup-detail-text appointment-popup-regarding-text">
                            ${escapeHtml(regarding)}
                          </span>
                        </div>
                      `
              : ""
            }
                </div>

                ${address
              ? `
                      <div class="appointment-popup-address-container">
                        <span
                          class="appointment-popup-address-text"
                          title="${escapeHtml(address)}"
                        >
                          ${escapeHtml(address)}
                        </span>
                      </div>
                    `
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
    if (!map || !appointmentSourceRef.current) return;

    const positions: atlas.data.Position[] = [];

    if (userLocation?.position) {
      positions.push(userLocation.position);
    }

    const shapes = appointmentSourceRef.current.getShapes();
    shapes.forEach((shape) => {
      const coords = shape.getCoordinates();
      if (coords) {
        positions.push(coords as atlas.data.Position);
      }
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
    if (!map || !popup || !appointmentSourceRef.current) return;

    if (!userLocation?.position) {
      return;
    }

    popup.close();

    appointmentSourceRef.current.clear();

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
    const features: atlas.data.Feature<atlas.data.Point, AppointmentProperties>[] = [];
    let markerCount = 0;

    for (const address of sortedAddresses) {
      const appts = addressMap.get(address);
      const position = geocodeMap.get(address);
      if (!position || !appts) continue;

      markerCount++;
      const isFirstMarker = markerCount === 1;

      const feature = new atlas.data.Feature(
        new atlas.data.Point(position),
        {
          appointments: appts,
          address: address,
          markerNumber: markerCount,
          isFirstMarker: isFirstMarker,
        }
      );

      features.push(feature);

      routePoints.push({
        position,
        address,
        appointmentId: appts[0].id,
        subject: appts[0].subject,
        scheduledstart: appts[0].scheduledstart,
      });
    }

    appointmentSourceRef.current.add(features);

    if (showRoute && routePoints.length > 0 && userLocation?.position) {
      await calculateAndDisplayRoute(userLocation.position, routePoints);
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
    if (!routeServiceRef.current || !routeSourceRef.current || !appointmentSourceRef.current) return;

    const filteredPoints = routePoints.filter(
      (point) => !positionsAreEqual(point.position, startPosition)
    );

    if (filteredPoints.length === 0) {
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

      let result: RouteResult | null = null;

      if (cached && Date.now() - cached.timestamp < ROUTE_CACHE_DURATION) {
        result = cached.result;
      } else {
        result = await routeServiceRef.current.calculateChronologicalRoute(
          startPosition,
          filteredPoints
        );

        if (result) {
          routeCacheRef.current.set(cacheKey, { result, timestamp: Date.now() });
        }
      }

      if (
        result &&
        result.routeCoordinates &&
        result.routeCoordinates.length > 0
      ) {
        setRouteData(result);
        displayRouteOnMap(result);
        const shapes = appointmentSourceRef.current.getShapes();
        const updatedFeatures: atlas.data.Feature<atlas.data.Point, AppointmentProperties>[] = [];

        shapes.forEach((shape) => {
          const coords = shape.getCoordinates();

          const feature = shape.toJson() as GeoJSON.Feature<GeoJSON.Point, AppointmentProperties>;
          const props = feature.properties;

          if (coords && props && props.markerNumber) {
            const routeLegIndex = props.markerNumber - 1;

            if (routeLegIndex >= 0 && routeLegIndex < result!.routeLegs.length) {
              const routeLegData = result!.routeLegs[routeLegIndex];

              const updatedFeature = new atlas.data.Feature(
                new atlas.data.Point(coords as atlas.data.Position),
                {
                  ...props,
                  routeLegDistance: routeLegData.distance,
                  routeLegDuration: routeLegData.duration,
                  routeLegFrom: routeLegData.from,
                  routeLegTo: routeLegData.to
                }
              );
              updatedFeatures.push(updatedFeature);
            } else {
              updatedFeatures.push(new atlas.data.Feature(
                new atlas.data.Point(coords as atlas.data.Position),
                props
              ));
            }
          }
        });

        if (updatedFeatures.length > 0) {
          appointmentSourceRef.current.clear();
          appointmentSourceRef.current.add(updatedFeatures);
        }
      } else {
        routeSourceRef.current.clear();
        setRouteData(null);
      }
    } catch (error) {
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