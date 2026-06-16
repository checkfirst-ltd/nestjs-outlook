import { Injectable, Logger, Inject, OnModuleInit } from '@nestjs/common';
import { AxiosResponse, InternalAxiosRequestConfig } from 'axios';
import { Client } from '@microsoft/microsoft-graph-client';
import type { Middleware } from '@microsoft/microsoft-graph-client';
import { MICROSOFT_CONFIG } from '../constants';
import { MicrosoftOutlookConfig } from '../interfaces/config/outlook-config.interface';
import { graphAxios } from './graph-axios';
import { ScenarioRegistry } from './scenario-registry';
import { MockRequest, MockResponse, MockRoute, MockScenario } from './interfaces';

// Import built-in scenarios (triggers auto-registration)
import './scenarios';

@Injectable()
export class MockGraphInterceptorService implements OnModuleInit {
  private readonly logger = new Logger(MockGraphInterceptorService.name);
  private scenario: MockScenario | undefined;
  private interceptorId: number | undefined;

  constructor(
    @Inject(MICROSOFT_CONFIG)
    private readonly config: MicrosoftOutlookConfig,
  ) {}

  onModuleInit(): void {
    if (!this.isEnabled()) return;

    const scenarioName = this.config.mock!.scenario;
    this.scenario = ScenarioRegistry.get(scenarioName);

    if (!this.scenario) {
      const available = ScenarioRegistry.list().join(', ');
      throw new Error(
        `Mock scenario "${scenarioName}" not found. Available: ${available}`,
      );
    }

    this.installAxiosInterceptor();
    this.logger.log(`Mock mode enabled with scenario: "${scenarioName}"`);
  }

  isEnabled(): boolean {
    return !!this.config.mock?.enabled;
  }

  /**
   * Creates a Graph Client that routes requests through the mock resolver
   * instead of making real HTTP calls.
   */
  createMockClient(): Client {
    const self = this;

    const mockMiddleware: Middleware = {
      execute: async (context) => {
        const url = context.request as string;
        const method = (context.options?.method || 'GET').toUpperCase();
        const body = context.options?.body
          ? JSON.parse(context.options.body as string)
          : undefined;

        const mockRequest: MockRequest = {
          method,
          url: url.startsWith('https://')
            ? url
            : `https://graph.microsoft.com/v1.0${url.startsWith('/') ? '' : '/'}${url}`,
          headers: {},
          body,
        };

        const mockResponse = await self.resolveResponse(mockRequest);

        if (mockResponse.delayMs) {
          await new Promise((r) => setTimeout(r, mockResponse.delayMs));
        }

        // Simulate error for non-2xx
        if (mockResponse.status >= 400) {
          const error = new Error(
            `Mock Graph API error: ${mockResponse.status}`,
          ) as Error & { statusCode: number; code: string; body: string };
          error.statusCode = mockResponse.status;
          const data = mockResponse.data as Record<string, unknown> | null;
          const errObj = data?.['error'] as Record<string, unknown> | undefined;
          error.code = (errObj?.['code'] as string) || 'MockError';
          error.body = JSON.stringify(mockResponse.data);
          throw error;
        }

        context.response = new Response(JSON.stringify(mockResponse.data), {
          status: mockResponse.status,
          headers: {
            'Content-Type': 'application/json',
            ...mockResponse.headers,
          },
        });
      },
    };

    return Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => 'mock-access-token',
      },
      middleware: mockMiddleware,
    });
  }

  async resolveResponse(request: MockRequest): Promise<MockResponse> {
    if (!this.scenario) {
      return { status: 200, data: {} };
    }

    const matchedRoute = this.findMatchingRoute(
      request,
      this.scenario.routes,
    );

    if (!matchedRoute) {
      this.logger.warn(
        `No mock route matched: ${request.method} ${request.url}. Returning default 200.`,
      );
      return { status: 200, data: {} };
    }

    return matchedRoute.handler(request);
  }

  private findMatchingRoute(
    request: MockRequest,
    routes: MockRoute[],
  ): MockRoute | undefined {
    return routes.find((route) => {
      if (route.method !== request.method) return false;

      if (typeof route.urlPattern === 'string') {
        return request.url.includes(route.urlPattern);
      }
      return route.urlPattern.test(request.url);
    });
  }

  private installAxiosInterceptor(): void {
    this.interceptorId = graphAxios.interceptors.request.use(
      async (config: InternalAxiosRequestConfig) => {
        if (!this.isGraphUrl(config.url)) return config;

        const method = (config.method || 'GET').toUpperCase();
        const mockRequest: MockRequest = {
          method,
          url: config.url!,
          headers: config.headers
            ? Object.fromEntries(
                Object.entries(config.headers).map(([k, v]) => [
                  k,
                  String(v),
                ]),
              )
            : {},
          body: config.data,
        };

        const mockResponse = await this.resolveResponse(mockRequest);

        if (mockResponse.delayMs) {
          await new Promise((r) => setTimeout(r, mockResponse.delayMs));
        }

        // Use custom adapter to short-circuit the real HTTP call
        config.adapter = async (): Promise<AxiosResponse> => {
          const response: AxiosResponse = {
            data: mockResponse.data,
            status: mockResponse.status,
            statusText: mockResponse.status < 400 ? 'OK' : 'Mock Error',
            headers: mockResponse.headers || {},
            config,
          };

          // Axios treats non-2xx as errors by default
          if (mockResponse.status >= 400) {
            const error = new Error(
              `Request failed with status code ${mockResponse.status}`,
            ) as Error & { response: AxiosResponse; config: InternalAxiosRequestConfig; isAxiosError: boolean };
            error.response = response;
            error.config = config;
            error.isAxiosError = true;
            throw error;
          }

          return response;
        };

        return config;
      },
    );
  }

  private isGraphUrl(url?: string): boolean {
    if (!url) return false;
    return (
      url.includes('graph.microsoft.com') ||
      url.includes('login.microsoftonline.com')
    );
  }
}
