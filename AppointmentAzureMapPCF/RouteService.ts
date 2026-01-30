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

interface RouteApiLegSummary {
  lengthInMeters: number;
  travelTimeInSeconds: number;
}

interface RouteApiLeg {
  summary: RouteApiLegSummary;
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

interface CacheEntry {
  result: RouteResult;
  timestamp: number;
}

export class AzureMapsRouteService {
  private azureMapsKey: string;
  private readonly baseUrl = "https://atlas.microsoft.com/route/directions/json";
  private cache = new Map<string, CacheEntry>();
  private readonly cacheTTL = 5 * 60 * 1000;
  private apiCallCount = 0;

  constructor(azureMapsKey: string) {
    this.azureMapsKey = azureMapsKey;
  }

  async calculateChronologicalRoute(
    startPosition: atlas.data.Position,
    routePoints: RoutePoint[]
  ): Promise<RouteResult | null> {
    if (!startPosition || routePoints.length === 0) {
      return null;
    }

    const sortedPoints = [...routePoints].sort(
      (a, b) => a.scheduledstart.getTime() - b.scheduledstart.getTime()
    );

    const cacheKey = this.generateCacheKey(startPosition, sortedPoints);
    const cached = this.getValidCacheEntry(cacheKey);
    if (cached) {
      return cached;
    }

    try {
      const waypoints = [startPosition, ...sortedPoints.map((p) => p.position)];
      const routeData = await this.getRoute(waypoints);

      if (!routeData?.routes?.[0]) {
        return null;
      }

      const result = this.buildRouteResult(startPosition, sortedPoints, routeData);

      this.cache.set(cacheKey, { result, timestamp: Date.now() });
      this.apiCallCount++;

      return result;
    } catch (error) {
      return null;
    }
  }

  private buildRouteResult(
    startPosition: atlas.data.Position,
    sortedPoints: RoutePoint[],
    routeData: RouteApiResponse
  ): RouteResult {
    const firstRoute = routeData.routes[0];

    return {
      routeCoordinates: this.extractRouteCoordinates(firstRoute.legs),
      totalDistance: firstRoute.summary.lengthInMeters,
      totalDuration: firstRoute.summary.travelTimeInSeconds,
      routeLegs: this.buildRouteLegDetails(sortedPoints, firstRoute.legs),
    };
  }

  private extractRouteCoordinates(legs: RouteApiLeg[]): atlas.data.Position[] {
    const coords: atlas.data.Position[] = [];

    for (const leg of legs) {
      for (const point of leg.points) {
        coords.push([point.longitude, point.latitude]);
      }
    }

    return coords;
  }

  private async getRoute(waypoints: atlas.data.Position[]): Promise<RouteApiResponse> {
    const waypointString = waypoints.map((pos) => `${pos[1]},${pos[0]}`).join(":");

    const params = new URLSearchParams({
      "api-version": "1.0",
      "subscription-key": this.azureMapsKey,
      query: waypointString,
      computeTravelTime: "true",
      travelMode: "car",
      traffic: "true",
    });

    const response = await fetch(`${this.baseUrl}?${params.toString()}`);

    if (!response.ok) {
      throw new Error(`Route API failed: ${response.statusText}`);
    }

    const data = await response.json();

    return data;
  }

  private buildRouteLegDetails(
    sortedPoints: RoutePoint[],
    legs: RouteApiLeg[]
  ): RouteLeg[] {
    const legDetails: RouteLeg[] = [];

    if (legs.length > 0) {
      const firstPoint = sortedPoints[0];
      const firstLeg = {
        from: "Your Location",
        to: firstPoint.subject || firstPoint.address || "First Appointment",
        distance: legs[0].summary.lengthInMeters,
        duration: legs[0].summary.travelTimeInSeconds,
        summary: AzureMapsRouteService.formatRouteSummary(
          legs[0].summary.lengthInMeters,
          legs[0].summary.travelTimeInSeconds
        ),
      };
      legDetails.push(firstLeg);
    }

    for (let i = 1; i < legs.length; i++) {
      const leg = legs[i];
      const fromPoint = sortedPoints[i - 1];
      const toPoint = sortedPoints[i];

      const routeLeg = {
        from: fromPoint.subject || fromPoint.address || `Stop ${i}`,
        to: toPoint.subject || toPoint.address || `Stop ${i + 1}`,
        distance: leg.summary.lengthInMeters,
        duration: leg.summary.travelTimeInSeconds,
        summary: AzureMapsRouteService.formatRouteSummary(
          leg.summary.lengthInMeters,
          leg.summary.travelTimeInSeconds
        ),
      };

      legDetails.push(routeLeg);
    }

    return legDetails;
  }

  private getValidCacheEntry(key: string): RouteResult | null {
    const entry = this.cache.get(key);
    if (!entry) return null;

    const isExpired = Date.now() - entry.timestamp > this.cacheTTL;
    if (isExpired) {
      this.cache.delete(key);
      return null;
    }

    return entry.result;
  }

  private pruneCache(): void {
    if (this.cache.size > 100) {
      this.cache.clear();
    }
  }

  private generateCacheKey(start: atlas.data.Position, points: RoutePoint[]): string {
    return `${start[0]},${start[1]}-${points.map((p) => `${p.appointmentId}`).join(",")}`;
  }

  private static formatRouteSummary(
    distanceInMeters: number,
    durationInSeconds: number
  ): string {
    const distanceInMiles = (distanceInMeters * 0.000621371).toFixed(1);
    const hours = Math.floor(durationInSeconds / 3600);
    const minutes = Math.floor((durationInSeconds % 3600) / 60);

    const durationStr = hours > 0 ? `${hours}h ${minutes}m` : `${minutes}m`;
    return `${distanceInMiles} mi • ${durationStr}`;
  }

  static createRouteGeoJSON(routeCoordinates: atlas.data.Position[]): GeoJSON.LineString {
    return {
      type: "LineString",
      coordinates: routeCoordinates,
    };
  }

  static formatTotalDuration(durationInSeconds: number): string {
    const hours = Math.floor(durationInSeconds / 3600);
    const minutes = Math.floor((durationInSeconds % 3600) / 60);

    return hours > 0 ? `${hours}h ${minutes}m` : `${minutes}m`;
  }

  static formatTotalDistance(distanceInMeters: number): string {
    const distanceInMiles = distanceInMeters * 0.000621371;
    return `${distanceInMiles.toFixed(1)} mi`;
  }
}