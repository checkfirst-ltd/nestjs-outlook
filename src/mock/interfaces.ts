export interface MockRequest {
  method: string;
  url: string;
  headers: Record<string, string>;
  body?: unknown;
}

export interface MockResponse {
  status: number;
  data: unknown;
  headers?: Record<string, string>;
  delayMs?: number;
}

export interface MockRoute {
  method: 'GET' | 'POST' | 'PATCH' | 'DELETE';
  urlPattern: string | RegExp;
  handler: (request: MockRequest) => MockResponse | Promise<MockResponse>;
}

export interface MockScenario {
  name: string;
  description: string;
  routes: MockRoute[];
}

