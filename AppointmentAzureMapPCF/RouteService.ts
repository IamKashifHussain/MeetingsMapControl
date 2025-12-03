import * as atlas from "azure-maps-control";

export interface RoutePoint {
  position: atlas.data.Position;
  address: string;
  appointmentId: string;
  subject: string;
  scheduledstart: Date;
}

export interface RouteResult {
  routeCoordinates: atlas.data.Position[];
  totalDistance: number;
  totalDuration: number;
  routeLegs: RouteLeg[];
}

export interface RouteLeg {
  from: string;
  to: string;
  distance: number;
  duration: number;
  summary: string;
}

interface RouteApiLegPoint {
  longitude: number;
  latitude: number;
}

interface RouteApiLeg {
  lengthInMeters: number;
  travelTimeInSeconds: number;
  points: RouteApiLegPoint[];
}

interface RouteApiSummary {
  lengthInMeters: number;
  travelTimeInSeconds: number;
}

interface RouteApiRoute {
  legs: RouteApiLeg[];
  summary: RouteApiSummary;
}

interface RouteApiResponse {
  routes: RouteApiRoute[];
}

export class AzureMapsRouteService {
  private azureMapsKey: string;

  constructor(azureMapsKey: string) {
    this.azureMapsKey = azureMapsKey;
  }

  /**
   * Calculates a route through multiple waypoints in chronological order
   */
  async calculateChronologicalRoute(
    startPosition: atlas.data.Position,
    routePoints: RoutePoint[]
  ): Promise<RouteResult | null> {
    if (!startPosition || routePoints.length === 0) {
      return null;
    }

    // Sort points by appointment scheduled start time
    const sortedPoints = [...routePoints].sort(
      (a, b) => a.scheduledstart.getTime() - b.scheduledstart.getTime()
    );

    try {
      // Build waypoints array: start position + all appointment locations
      const waypoints = [
        startPosition,
        ...sortedPoints.map((point) => point.position),
      ];

      // Calculate the route using Azure Maps Directions API
      const routeData = await this.getRoute(waypoints);

      if (!routeData) {
        return null;
      }

      // Extract route legs with detailed information
      const routeLegs = this.buildRouteLegDetails(
        startPosition,
        sortedPoints,
        routeData
      );

      return {
        routeCoordinates: routeData.routes[0].legs.reduce(
          (acc: atlas.data.Position[], leg: RouteApiLeg) => [
            ...acc,
            ...leg.points.map((p: RouteApiLegPoint) => [p.longitude, p.latitude] as atlas.data.Position),
          ],
          [] as atlas.data.Position[]
        ),
        totalDistance: routeData.routes[0].summary.lengthInMeters,
        totalDuration: routeData.routes[0].summary.travelTimeInSeconds,
        routeLegs: routeLegs,
      };
    } catch (error) {
      console.error("Route calculation failed:", error);
      return null;
    }
  }

  /**
   * Fetches route data from Azure Maps Directions API
   */
  private async getRoute(
    waypoints: atlas.data.Position[]
  ): Promise<RouteApiResponse> {
    const waypointString = waypoints
      .map((pos) => `${pos[1]},${pos[0]}`)
      .join(":");

    const url =
      `https://atlas.microsoft.com/route/directions/json?` +
      `api-version=1.0&` +
      `subscription-key=${this.azureMapsKey}&` +
      `query=${waypointString}&` +
      `computeTravelTime=true&` +
      `travelMode=car&` +
      `traffic=true`;

    const response = await fetch(url);

    if (!response.ok) {
      throw new Error(`Route API failed: ${response.statusText}`);
    }

    return await response.json() as RouteApiResponse;
  }

  /**
   * Builds detailed leg information from route data
   */
  private buildRouteLegDetails(
    startPosition: atlas.data.Position,
    sortedPoints: RoutePoint[],
    routeData: RouteApiResponse
  ): RouteLeg[] {
    const legs: RouteLeg[] = [];
    const routeLegs = routeData.routes[0].legs;

    // First leg: from user location to first appointment
    if (routeLegs.length > 0) {
      legs.push({
        from: "Your Location",
        to: sortedPoints[0]?.subject || sortedPoints[0]?.address || "First Appointment",
        distance: routeLegs[0].lengthInMeters,
        duration: routeLegs[0].travelTimeInSeconds,
        summary: this.formatRouteSummary(
          routeLegs[0].lengthInMeters,
          routeLegs[0].travelTimeInSeconds
        ),
      });
    }

    // Subsequent legs: between appointments
    for (let i = 1; i < routeLegs.length; i++) {
      const leg = routeLegs[i];
      const fromPoint = sortedPoints[i - 1];
      const toPoint = sortedPoints[i];

      legs.push({
        from: fromPoint?.subject || fromPoint?.address || `Stop ${i}`,
        to: toPoint?.subject || toPoint?.address || `Stop ${i + 1}`,
        distance: leg.lengthInMeters,
        duration: leg.travelTimeInSeconds,
        summary: this.formatRouteSummary(
          leg.lengthInMeters,
          leg.travelTimeInSeconds
        ),
      });
    }

    return legs;
  }

  /**
   * Formats distance and duration into human-readable string
   */
  private formatRouteSummary(distanceInMeters: number, durationInSeconds: number): string {
    const distanceInKm = (distanceInMeters / 1000).toFixed(1);
    const hours = Math.floor(durationInSeconds / 3600);
    const minutes = Math.floor((durationInSeconds % 3600) / 60);

    let durationStr = "";
    if (hours > 0) {
      durationStr = `${hours}h ${minutes}m`;
    } else {
      durationStr = `${minutes}m`;
    }

    return `${distanceInKm} km • ${durationStr}`;
  }

  /**
   * Converts route into GeoJSON LineString for visualization
   */
  static createRouteGeoJSON(
    routeCoordinates: atlas.data.Position[]
  ): GeoJSON.LineString {
    return {
      type: "LineString",
      coordinates: routeCoordinates.map((pos) => [pos[0], pos[1]]),
    };
  }

  /**
   * Formats total duration
   */
  static formatTotalDuration(durationInSeconds: number): string {
    const hours = Math.floor(durationInSeconds / 3600);
    const minutes = Math.floor((durationInSeconds % 3600) / 60);

    if (hours > 0) {
      return `${hours}h ${minutes}m`;
    }
    return `${minutes}m`;
  }

  /**
   * Formats total distance
   */
  static formatTotalDistance(distanceInMeters: number): string {
    if (distanceInMeters >= 1000) {
      return `${(distanceInMeters / 1000).toFixed(1)} km`;
    }
    return `${Math.round(distanceInMeters)} m`;
  }
}