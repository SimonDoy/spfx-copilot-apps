import type { IWeather } from '../models/myDay';

/**
 * Live weather for the My Day hero.
 *
 * Uses **Open-Meteo** — free, keyless and CORS-enabled — so the sample stays
 * deployable by anyone without provisioning an API key. Three anonymous GETs:
 * geocode the city, read current conditions, read the European air quality
 * index. Any failure resolves to `undefined` and the caller falls back to the
 * mock weather, so the hero never breaks.
 */

/** Narrow port for anonymous JSON GETs, so this file stays free of SPFx types. */
export type JsonGet = (url: string) => Promise<unknown>;

export interface IWeatherService {
  getWeather(city: string, country: string): Promise<IWeather | undefined>;
}

const GEOCODE_URL = 'https://geocoding-api.open-meteo.com/v1/search';
const FORECAST_URL = 'https://api.open-meteo.com/v1/forecast';
const AIR_QUALITY_URL = 'https://air-quality-api.open-meteo.com/v1/air-quality';

/** WMO weather interpretation codes → the short label shown in the hero. */
const WMO_CONDITIONS: { [code: number]: string } = {
  0: 'Clear',
  1: 'Mostly clear',
  2: 'Partly cloudy',
  3: 'Cloudy',
  45: 'Fog',
  48: 'Freezing fog',
  51: 'Light drizzle',
  53: 'Drizzle',
  55: 'Heavy drizzle',
  56: 'Freezing drizzle',
  57: 'Freezing drizzle',
  61: 'Light rain',
  63: 'Rain',
  65: 'Heavy rain',
  66: 'Freezing rain',
  67: 'Freezing rain',
  71: 'Light snow',
  73: 'Snow',
  75: 'Heavy snow',
  77: 'Snow grains',
  80: 'Rain showers',
  81: 'Rain showers',
  82: 'Heavy rain showers',
  85: 'Snow showers',
  86: 'Heavy snow showers',
  95: 'Thunderstorm',
  96: 'Thunderstorm with hail',
  99: 'Thunderstorm with hail'
};

/** European AQI bands (0-20 good … 100+ extremely poor). */
const aqiLabel = (index: number): string => {
  if (index <= 20) return 'Good';
  if (index <= 40) return 'Fair';
  if (index <= 60) return 'Moderate';
  if (index <= 80) return 'Poor';
  if (index <= 100) return 'Very poor';
  return 'Extremely poor';
};

interface IGeocodeResult {
  results?: { latitude: number; longitude: number; name: string; country?: string }[];
}

interface IForecastResult {
  current?: { temperature_2m?: number; weather_code?: number };
}

interface IAirQualityResult {
  current?: { european_aqi?: number };
}

export class OpenMeteoWeatherService implements IWeatherService {
  private readonly getJson: JsonGet;

  public constructor(getJson: JsonGet) {
    this.getJson = getJson;
  }

  public async getWeather(city: string, country: string): Promise<IWeather | undefined> {
    if (!city) {
      return undefined;
    }

    try {
      const query = encodeURIComponent(country ? `${city}, ${country}` : city);
      const geo = (await this.getJson(
        `${GEOCODE_URL}?name=${query}&count=1&format=json`
      )) as IGeocodeResult;

      const place = geo.results?.[0];
      if (!place) {
        return undefined;
      }

      const coords = `latitude=${place.latitude}&longitude=${place.longitude}`;

      const [forecast, air] = await Promise.all([
        this.getJson(`${FORECAST_URL}?${coords}&current=temperature_2m,weather_code`) as Promise<IForecastResult>,
        // Air quality is a bonus signal; a failure here must not lose the temperature.
        (this.getJson(`${AIR_QUALITY_URL}?${coords}&current=european_aqi`) as Promise<IAirQualityResult>).catch(
          () => ({}) as IAirQualityResult
        )
      ]);

      const temperatureC = forecast.current?.temperature_2m;
      if (typeof temperatureC !== 'number') {
        return undefined;
      }

      const code = forecast.current?.weather_code;
      const aqi = Math.round(air.current?.european_aqi ?? 0);

      return {
        temperatureC: Math.round(temperatureC),
        temperatureF: Math.round(temperatureC * 1.8 + 32),
        condition: (code !== undefined && WMO_CONDITIONS[code]) || 'Unknown',
        location: place.country ? `${place.name}, ${place.country}` : place.name,
        airQualityIndex: aqi,
        airQualityLabel: aqiLabel(aqi)
      };
    } catch {
      return undefined;
    }
  }
}
