import { Test, TestingModule } from '@nestjs/testing';
import { EventEmitter2 } from '@nestjs/event-emitter';
import { getRepositoryToken } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import axios from 'axios';
import { TenantCalendarService } from './tenant-calendar.service';
import { AppOnlyAuthService } from '../auth/app-only-auth.service';
import { MicrosoftTenantUser } from '../../entities/microsoft-tenant-user.entity';
import { MICROSOFT_CONFIG } from '../../constants';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';

jest.mock('axios');
const mockedAxios = axios as jest.Mocked<typeof axios>;

describe('TenantCalendarService', () => {
  let service: TenantCalendarService;
  let appOnlyAuthService: jest.Mocked<AppOnlyAuthService>;
  let tenantUserRepository: jest.Mocked<Repository<MicrosoftTenantUser>>;
  let eventEmitter: jest.Mocked<EventEmitter2>;

  const mockConfig: MicrosoftOutlookConfig = {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    redirectPath: '/auth/callback',
    backendBaseUrl: 'https://test.example.com',
  };

  const mockTenantId = 'tenant-123-guid';
  const mockMicrosoftUserId = 'user-456-guid';
  const mockCalendarId = 'calendar-789-guid';
  const mockAccessToken = 'mock-access-token';

  beforeEach(async () => {
    const mockAppOnlyAuthService = {
      getAccessToken: jest.fn().mockResolvedValue(mockAccessToken),
      isEnabled: jest.fn().mockReturnValue(true),
    };

    const mockTenantUserRepository = {
      update: jest.fn().mockResolvedValue({ affected: 1 }),
      findOne: jest.fn(),
    };

    const mockEventEmitter = {
      emit: jest.fn(),
    };

    const module: TestingModule = await Test.createTestingModule({
      providers: [
        TenantCalendarService,
        {
          provide: AppOnlyAuthService,
          useValue: mockAppOnlyAuthService,
        },
        {
          provide: getRepositoryToken(MicrosoftTenantUser),
          useValue: mockTenantUserRepository,
        },
        {
          provide: EventEmitter2,
          useValue: mockEventEmitter,
        },
        {
          provide: MICROSOFT_CONFIG,
          useValue: mockConfig,
        },
      ],
    }).compile();

    service = module.get<TenantCalendarService>(TenantCalendarService);
    appOnlyAuthService = module.get(AppOnlyAuthService);
    tenantUserRepository = module.get(getRepositoryToken(MicrosoftTenantUser));
    eventEmitter = module.get(EventEmitter2);

    jest.clearAllMocks();
  });

  afterEach(() => {
    service.clearCache();
  });

  describe('getDefaultCalendarId', () => {
    it('should use /users/{id}/calendar endpoint', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { id: mockCalendarId, name: 'Calendar' },
      });

      await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);

      expect(mockedAxios.get).toHaveBeenCalledWith(
        `https://graph.microsoft.com/v1.0/users/${mockMicrosoftUserId}/calendar`,
        expect.objectContaining({
          headers: expect.objectContaining({
            Authorization: `Bearer ${mockAccessToken}`,
          }),
        })
      );
    });

    it('should include IdType="ImmutableId" header', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { id: mockCalendarId, name: 'Calendar' },
      });

      await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);

      expect(mockedAxios.get).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          headers: expect.objectContaining({
            Prefer: expect.stringContaining('IdType="ImmutableId"'),
          }),
        })
      );
    });

    it('should include outlook.timezone="UTC" header', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { id: mockCalendarId, name: 'Calendar' },
      });

      await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);

      expect(mockedAxios.get).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          headers: expect.objectContaining({
            Prefer: expect.stringContaining('outlook.timezone="UTC"'),
          }),
        })
      );
    });

    it('should cache calendar IDs', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { id: mockCalendarId, name: 'Calendar' },
      });

      // First call - should hit API
      const result1 = await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);
      expect(result1).toBe(mockCalendarId);
      expect(mockedAxios.get).toHaveBeenCalledTimes(1);

      // Second call - should use cache
      const result2 = await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);
      expect(result2).toBe(mockCalendarId);
      expect(mockedAxios.get).toHaveBeenCalledTimes(1); // Still 1, no additional call
    });

    it('should persist calendar ID to database', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { id: mockCalendarId, name: 'Calendar' },
      });

      await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);

      expect(tenantUserRepository.update).toHaveBeenCalledWith(
        { microsoftUserId: mockMicrosoftUserId, isActive: true },
        { defaultCalendarId: mockCalendarId }
      );
    });
  });

  describe('createEvent', () => {
    const mockEvent = {
      subject: 'Test Meeting',
      start: { dateTime: '2024-01-15T10:00:00', timeZone: 'UTC' },
      end: { dateTime: '2024-01-15T11:00:00', timeZone: 'UTC' },
    };

    const mockCreatedEvent = {
      id: 'event-123',
      ...mockEvent,
    };

    it('should use /users/{id}/calendars/{calendarId}/events endpoint', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: mockCreatedEvent,
      });

      await service.createEvent(mockEvent, mockTenantId, mockMicrosoftUserId, mockCalendarId);

      expect(mockedAxios.post).toHaveBeenCalledWith(
        `https://graph.microsoft.com/v1.0/users/${mockMicrosoftUserId}/calendars/${mockCalendarId}/events`,
        mockEvent,
        expect.objectContaining({
          headers: expect.objectContaining({
            Authorization: `Bearer ${mockAccessToken}`,
          }),
        })
      );
    });

    it('should include IdType="ImmutableId" header on create', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: mockCreatedEvent,
      });

      await service.createEvent(mockEvent, mockTenantId, mockMicrosoftUserId, mockCalendarId);

      expect(mockedAxios.post).toHaveBeenCalledWith(
        expect.any(String),
        expect.any(Object),
        expect.objectContaining({
          headers: expect.objectContaining({
            Prefer: expect.stringContaining('IdType="ImmutableId"'),
          }),
        })
      );
    });

    it('should return created event', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: mockCreatedEvent,
      });

      const result = await service.createEvent(mockEvent, mockTenantId, mockMicrosoftUserId, mockCalendarId);

      expect(result.event).toEqual(mockCreatedEvent);
    });
  });

  describe('getEventById', () => {
    const mockEvent = {
      id: 'event-123',
      subject: 'Test Meeting',
    };

    it('should use /users/{id}/events/{eventId} endpoint', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: mockEvent,
      });

      await service.getEventById(mockTenantId, mockMicrosoftUserId, 'event-123');

      expect(mockedAxios.get).toHaveBeenCalledWith(
        `https://graph.microsoft.com/v1.0/users/${mockMicrosoftUserId}/events/event-123`,
        expect.objectContaining({
          headers: expect.objectContaining({
            Authorization: `Bearer ${mockAccessToken}`,
          }),
        })
      );
    });

    it('should return null for 404 response', async () => {
      const error = new Error('Not Found');
      (error as unknown as { response: { status: number } }).response = { status: 404 };
      mockedAxios.get.mockRejectedValueOnce(error);

      // The executeGraphApiCall utility handles 404 and returns null
      mockedAxios.get.mockResolvedValueOnce(null);

      const result = await service.getEventById(mockTenantId, mockMicrosoftUserId, 'deleted-event');

      expect(result).toBeNull();
    });
  });

  describe('updateEvent', () => {
    const mockUpdates = {
      subject: 'Updated Meeting',
    };

    const mockUpdatedEvent = {
      id: 'event-123',
      subject: 'Updated Meeting',
    };

    it('should use PATCH on /users/{id}/calendars/{calendarId}/events/{eventId}', async () => {
      mockedAxios.patch.mockResolvedValueOnce({
        data: mockUpdatedEvent,
      });

      await service.updateEvent('event-123', mockUpdates, mockTenantId, mockMicrosoftUserId, mockCalendarId);

      expect(mockedAxios.patch).toHaveBeenCalledWith(
        `https://graph.microsoft.com/v1.0/users/${mockMicrosoftUserId}/calendars/${mockCalendarId}/events/event-123`,
        mockUpdates,
        expect.objectContaining({
          headers: expect.objectContaining({
            Authorization: `Bearer ${mockAccessToken}`,
          }),
        })
      );
    });
  });

  describe('deleteEvent', () => {
    it('should use DELETE on /users/{id}/calendars/{calendarId}/events/{eventId}', async () => {
      mockedAxios.delete.mockResolvedValueOnce({
        status: 204,
      });

      await service.deleteEvent(mockTenantId, mockMicrosoftUserId, 'event-123', mockCalendarId);

      expect(mockedAxios.delete).toHaveBeenCalledWith(
        `https://graph.microsoft.com/v1.0/users/${mockMicrosoftUserId}/calendars/${mockCalendarId}/events/event-123`,
        expect.objectContaining({
          headers: expect.objectContaining({
            Authorization: `Bearer ${mockAccessToken}`,
          }),
        })
      );
    });
  });

  describe('createBatchEvents', () => {
    it('should use $batch endpoint with /users/{id} paths', async () => {
      const events = [
        { subject: 'Meeting 1' },
        { subject: 'Meeting 2' },
      ];

      mockedAxios.post.mockResolvedValueOnce({
        data: {
          responses: [
            { id: '0', status: 201, body: { id: 'event-1', subject: 'Meeting 1' } },
            { id: '1', status: 201, body: { id: 'event-2', subject: 'Meeting 2' } },
          ],
        },
      });

      await service.createBatchEvents(events, mockTenantId, mockMicrosoftUserId, mockCalendarId);

      expect(mockedAxios.post).toHaveBeenCalledWith(
        'https://graph.microsoft.com/v1.0/$batch',
        expect.objectContaining({
          requests: expect.arrayContaining([
            expect.objectContaining({
              method: 'POST',
              url: `/users/${mockMicrosoftUserId}/calendars/${mockCalendarId}/events`,
            }),
          ]),
        }),
        expect.any(Object)
      );
    });
  });

  describe('token acquisition', () => {
    it('should get access token from AppOnlyAuthService', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { id: mockCalendarId },
      });

      await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);

      expect(appOnlyAuthService.getAccessToken).toHaveBeenCalledWith(mockTenantId);
    });
  });
});
