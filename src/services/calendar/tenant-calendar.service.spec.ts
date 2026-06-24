import { Test, TestingModule } from '@nestjs/testing';
import { getRepositoryToken } from '@nestjs/typeorm';
import { EventEmitter2 } from '@nestjs/event-emitter';
import { Repository } from 'typeorm';
import axios from 'axios';
import { TenantCalendarService } from '../tenant/tenant-calendar.service';
import { AppOnlyAuthService } from '../auth/app-only-auth.service';
import { MicrosoftTenantUser } from '../../entities/microsoft-tenant-user.entity';
import { MICROSOFT_CONFIG } from '../../constants';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';
import { Event } from '../../types';

jest.mock('axios');
const mockedAxios = axios as jest.Mocked<typeof axios>;

describe('TenantCalendarService', () => {
  let service: TenantCalendarService;
  let appOnlyAuthService: jest.Mocked<AppOnlyAuthService>;
  let tenantUserRepository: jest.Mocked<Repository<MicrosoftTenantUser>>;
  let eventEmitter: jest.Mocked<EventEmitter2>;

  const mockTenantId = '12345678-1234-1234-1234-123456789abc';
  const mockMicrosoftUserId = 'user-guid-12345';
  const mockCalendarId = 'calendar-guid-12345';
  const mockAccessToken = 'mock-access-token';
  const mockEventId = 'event-guid-12345';

  const mockConfig: MicrosoftOutlookConfig = {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    redirectPath: '/auth/callback',
    backendBaseUrl: 'https://api.example.com',
    appOnly: {
      enabled: true,
      tenantId: mockTenantId,
    },
  };

  const mockEvent: Event = {
    id: mockEventId,
    subject: 'Test Event',
    start: { dateTime: '2026-06-25T10:00:00', timeZone: 'UTC' },
    end: { dateTime: '2026-06-25T11:00:00', timeZone: 'UTC' },
    bodyPreview: 'Test event body',
  };

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

  describe('endpoint patterns', () => {
    it('should use /users/{id}/calendar endpoint pattern for getDefaultCalendarId', async () => {
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

    it('should use /users/{id}/calendars/{calendarId}/events for createEvent', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: mockEvent,
      });

      await service.createEvent(
        { subject: 'Test Event' },
        mockTenantId,
        mockMicrosoftUserId,
        mockCalendarId
      );

      expect(mockedAxios.post).toHaveBeenCalledWith(
        `https://graph.microsoft.com/v1.0/users/${mockMicrosoftUserId}/calendars/${mockCalendarId}/events`,
        expect.any(Object),
        expect.any(Object)
      );
    });

    it('should use /users/{id}/events/{eventId} for getEventById', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: mockEvent,
      });

      await service.getEventById(mockTenantId, mockMicrosoftUserId, mockEventId);

      expect(mockedAxios.get).toHaveBeenCalledWith(
        `https://graph.microsoft.com/v1.0/users/${mockMicrosoftUserId}/events/${mockEventId}`,
        expect.any(Object)
      );
    });

    it('should support Microsoft Graph user ID (GUID) as identifier', async () => {
      const guidUserId = '87d349ed-44d7-43e1-9a83-5f2406dee5bd';
      mockedAxios.get.mockResolvedValueOnce({
        data: { id: mockCalendarId },
      });

      await service.getDefaultCalendarId(mockTenantId, guidUserId);

      expect(mockedAxios.get).toHaveBeenCalledWith(
        `https://graph.microsoft.com/v1.0/users/${guidUserId}/calendar`,
        expect.any(Object)
      );
    });
  });

  describe('immutable ID header', () => {
    it('should include Prefer: IdType="ImmutableId" header in requests', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { id: mockCalendarId },
      });

      await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);

      expect(mockedAxios.get).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          headers: expect.objectContaining({
            'Prefer': 'IdType="ImmutableId", outlook.timezone="UTC"',
          }),
        })
      );
    });

    it('should include IdType="ImmutableId" in batch request headers', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: {
          responses: [{ id: '0', status: 201, body: mockEvent }],
        },
      });

      await service.createBatchEvents(
        [{ subject: 'Test' }],
        mockTenantId,
        mockMicrosoftUserId,
        mockCalendarId
      );

      const batchCall = mockedAxios.post.mock.calls[0];
      const batchPayload = batchCall[1] as { requests: Array<{ headers: Record<string, string> }> };
      expect(batchPayload.requests[0].headers['Prefer']).toContain('IdType="ImmutableId"');
    });
  });

  describe('event creation', () => {
    it('should create event on specified user calendar', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: mockEvent,
      });

      const result = await service.createEvent(
        { subject: 'Test Event' },
        mockTenantId,
        mockMicrosoftUserId,
        mockCalendarId
      );

      expect(result.event).toEqual(mockEvent);
      expect(appOnlyAuthService.getAccessToken).toHaveBeenCalledWith(mockTenantId);
    });

    it('should support creating events with attendees', async () => {
      const eventWithAttendees = {
        subject: 'Meeting',
        attendees: [
          { emailAddress: { address: 'attendee@contoso.com' }, type: 'required' },
        ],
      };

      mockedAxios.post.mockResolvedValueOnce({
        data: { ...mockEvent, ...eventWithAttendees },
      });

      await service.createEvent(
        eventWithAttendees,
        mockTenantId,
        mockMicrosoftUserId,
        mockCalendarId
      );

      expect(mockedAxios.post).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          attendees: expect.arrayContaining([
            expect.objectContaining({
              emailAddress: { address: 'attendee@contoso.com' },
            }),
          ]),
        }),
        expect.any(Object)
      );
    });

    it('should throw error when event creation fails', async () => {
      mockedAxios.post.mockRejectedValueOnce(new Error('Network error'));

      await expect(
        service.createEvent(
          { subject: 'Test Event' },
          mockTenantId,
          mockMicrosoftUserId,
          mockCalendarId
        )
      ).rejects.toThrow('Failed to create calendar event');
    });
  });

  describe('event modification', () => {
    it('should update event with partial payload using PATCH', async () => {
      const updatedEvent = { ...mockEvent, subject: 'Updated Subject' };
      mockedAxios.patch.mockResolvedValueOnce({
        data: updatedEvent,
      });

      const result = await service.updateEvent(
        mockEventId,
        { subject: 'Updated Subject' },
        mockTenantId,
        mockMicrosoftUserId,
        mockCalendarId
      );

      expect(mockedAxios.patch).toHaveBeenCalledWith(
        `https://graph.microsoft.com/v1.0/users/${mockMicrosoftUserId}/calendars/${mockCalendarId}/events/${mockEventId}`,
        { subject: 'Updated Subject' },
        expect.any(Object)
      );
      expect(result.event.subject).toBe('Updated Subject');
    });

    it('should delete event by ID', async () => {
      mockedAxios.delete.mockResolvedValueOnce({ status: 204 });

      await service.deleteEvent(
        mockTenantId,
        mockMicrosoftUserId,
        mockEventId,
        mockCalendarId
      );

      expect(mockedAxios.delete).toHaveBeenCalledWith(
        `https://graph.microsoft.com/v1.0/users/${mockMicrosoftUserId}/calendars/${mockCalendarId}/events/${mockEventId}`,
        expect.any(Object)
      );
    });
  });

  describe('batch operations', () => {
    it('should create batch events using $batch endpoint', async () => {
      const events = [{ subject: 'Event 1' }, { subject: 'Event 2' }];
      mockedAxios.post.mockResolvedValueOnce({
        data: {
          responses: [
            { id: '0', status: 201, body: { ...mockEvent, subject: 'Event 1' } },
            { id: '1', status: 201, body: { ...mockEvent, subject: 'Event 2' } },
          ],
        },
      });

      const results = await service.createBatchEvents(
        events,
        mockTenantId,
        mockMicrosoftUserId,
        mockCalendarId
      );

      expect(mockedAxios.post).toHaveBeenCalledWith(
        'https://graph.microsoft.com/v1.0/$batch',
        expect.objectContaining({
          requests: expect.arrayContaining([
            expect.objectContaining({ method: 'POST', id: '0' }),
            expect.objectContaining({ method: 'POST', id: '1' }),
          ]),
        }),
        expect.any(Object)
      );
      expect(results).toHaveLength(2);
      expect(results.every(r => r.success)).toBe(true);
    });

    it('should handle partial batch failures', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: {
          responses: [
            { id: '0', status: 201, body: mockEvent },
            { id: '1', status: 400, body: { error: { message: 'Invalid event' } } },
          ],
        },
      });

      const results = await service.createBatchEvents(
        [{ subject: 'Event 1' }, { subject: 'Event 2' }],
        mockTenantId,
        mockMicrosoftUserId,
        mockCalendarId
      );

      expect(results[0].success).toBe(true);
      expect(results[1].success).toBe(false);
      expect(results[1].error).toContain('HTTP 400');
    });

    it('should update batch events using PATCH method', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: {
          responses: [
            { id: '0', status: 200, body: mockEvent },
            { id: '1', status: 200, body: mockEvent },
          ],
        },
      });

      const results = await service.updateBatchEvents(
        [
          { eventId: 'event-1', updates: { subject: 'Updated 1' } },
          { eventId: 'event-2', updates: { subject: 'Updated 2' } },
        ],
        mockTenantId,
        mockMicrosoftUserId,
        mockCalendarId
      );

      const batchCall = mockedAxios.post.mock.calls[0];
      const batchPayload = batchCall[1] as { requests: Array<{ method: string }> };
      expect(batchPayload.requests.every(r => r.method === 'PATCH')).toBe(true);
      expect(results.every(r => r.success)).toBe(true);
    });

    it('should delete batch events with correct method', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: {
          responses: [
            { id: '0', status: 204 },
            { id: '1', status: 204 },
          ],
        },
      });

      const results = await service.deleteBatchEvents(
        ['event-1', 'event-2'],
        mockTenantId,
        mockMicrosoftUserId,
        mockCalendarId
      );

      const batchCall = mockedAxios.post.mock.calls[0];
      const batchPayload = batchCall[1] as { requests: Array<{ method: string }> };
      expect(batchPayload.requests.every(r => r.method === 'DELETE')).toBe(true);
      expect(results.every(r => r.success)).toBe(true);
    });

    it('should treat 404 as success for batch delete (already deleted)', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: {
          responses: [
            { id: '0', status: 404 },
          ],
        },
      });

      const results = await service.deleteBatchEvents(
        ['already-deleted-event'],
        mockTenantId,
        mockMicrosoftUserId,
        mockCalendarId
      );

      expect(results[0].success).toBe(true);
    });

    it('should respect batch size limit of 20', async () => {
      const manyEvents = Array.from({ length: 25 }, (_, i) => ({ subject: `Event ${i}` }));
      mockedAxios.post
        .mockResolvedValueOnce({
          data: {
            responses: Array.from({ length: 20 }, (_, i) => ({
              id: String(i),
              status: 201,
              body: { ...mockEvent, id: `event-${i}` },
            })),
          },
        })
        .mockResolvedValueOnce({
          data: {
            responses: Array.from({ length: 5 }, (_, i) => ({
              id: String(i),
              status: 201,
              body: { ...mockEvent, id: `event-${20 + i}` },
            })),
          },
        });

      const results = await service.createBatchEvents(
        manyEvents,
        mockTenantId,
        mockMicrosoftUserId,
        mockCalendarId
      );

      expect(mockedAxios.post).toHaveBeenCalledTimes(2);
      expect(results).toHaveLength(25);
    });
  });

  describe('token acquisition', () => {
    it('should acquire app-only token for Graph calls', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { id: mockCalendarId },
      });

      await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);

      expect(appOnlyAuthService.getAccessToken).toHaveBeenCalledWith(mockTenantId);
    });

    it('should include Bearer token in Authorization header', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: mockEvent,
      });

      await service.getEventById(mockTenantId, mockMicrosoftUserId, mockEventId);

      expect(mockedAxios.get).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          headers: expect.objectContaining({
            Authorization: `Bearer ${mockAccessToken}`,
          }),
        })
      );
    });
  });

  describe('caching', () => {
    it('should cache calendar ID after first fetch', async () => {
      mockedAxios.get.mockResolvedValue({
        data: { id: mockCalendarId },
      });

      // First call - should fetch from API
      await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);

      // Second call - should use cache
      await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);

      // Should only call API once
      expect(mockedAxios.get).toHaveBeenCalledTimes(1);
    });

    it('should update database with calendar ID', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { id: mockCalendarId },
      });

      await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);

      expect(tenantUserRepository.update).toHaveBeenCalledWith(
        { microsoftUserId: mockMicrosoftUserId, isActive: true },
        { defaultCalendarId: mockCalendarId }
      );
    });

    it('should clear cache when clearCache is called', async () => {
      mockedAxios.get.mockResolvedValue({
        data: { id: mockCalendarId },
      });

      // First call
      await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);

      // Clear cache
      service.clearCache();

      // Second call - should fetch from API again
      await service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId);

      expect(mockedAxios.get).toHaveBeenCalledTimes(2);
    });
  });

  describe('error handling', () => {
    it('should throw descriptive error when calendar fetch fails', async () => {
      mockedAxios.get.mockRejectedValueOnce(new Error('Network timeout'));

      await expect(
        service.getDefaultCalendarId(mockTenantId, mockMicrosoftUserId)
      ).rejects.toThrow('Failed to get calendar ID from Microsoft');
    });

    it('should return null for 404 on getEventById', async () => {
      mockedAxios.get.mockResolvedValueOnce(null);

      const result = await service.getEventById(
        mockTenantId,
        mockMicrosoftUserId,
        'non-existent-event'
      );

      expect(result).toBeNull();
    });

    it('should throw error when event update fails', async () => {
      mockedAxios.patch.mockRejectedValueOnce(new Error('Update failed'));

      await expect(
        service.updateEvent(
          mockEventId,
          { subject: 'Updated' },
          mockTenantId,
          mockMicrosoftUserId,
          mockCalendarId
        )
      ).rejects.toThrow('Failed to update calendar event');
    });

    it('should throw error when event deletion fails', async () => {
      mockedAxios.delete.mockRejectedValueOnce(new Error('Delete failed'));

      await expect(
        service.deleteEvent(
          mockTenantId,
          mockMicrosoftUserId,
          mockEventId,
          mockCalendarId
        )
      ).rejects.toThrow('Failed to delete calendar event');
    });
  });

  describe('streaming events', () => {
    it('should stream events using calendarView endpoint', async () => {
      const events = [mockEvent, { ...mockEvent, id: 'event-2' }];
      mockedAxios.get.mockResolvedValueOnce({
        data: {
          value: events,
          '@odata.nextLink': undefined,
        },
      });

      const chunks: Event[][] = [];
      for await (const chunk of service.streamEvents(mockTenantId, mockMicrosoftUserId)) {
        chunks.push(chunk);
      }

      expect(chunks.flat()).toHaveLength(2);
      expect(mockedAxios.get).toHaveBeenCalledWith(
        expect.stringContaining('/calendarView'),
        expect.any(Object)
      );
    });

    it('should handle pagination with nextLink', async () => {
      mockedAxios.get
        .mockResolvedValueOnce({
          data: {
            value: [mockEvent],
            '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next-page',
          },
        })
        .mockResolvedValueOnce({
          data: {
            value: [{ ...mockEvent, id: 'event-2' }],
          },
        });

      const chunks: Event[][] = [];
      for await (const chunk of service.streamEvents(mockTenantId, mockMicrosoftUserId)) {
        chunks.push(chunk);
      }

      expect(mockedAxios.get).toHaveBeenCalledTimes(2);
      expect(chunks.flat()).toHaveLength(2);
    });

    it('should emit IMPORT_COMPLETED event after streaming', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockEvent] },
      });

      const chunks: Event[][] = [];
      for await (const chunk of service.streamEvents(mockTenantId, mockMicrosoftUserId)) {
        chunks.push(chunk);
      }

      expect(eventEmitter.emit).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          tenantId: mockTenantId,
          microsoftUserId: mockMicrosoftUserId,
          isTenantWide: true,
        })
      );
    });
  });

  describe('getEventsBatch', () => {
    it('should fetch multiple events in a single batch request', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: {
          responses: [
            { id: '0', status: 200, body: mockEvent },
            { id: '1', status: 200, body: { ...mockEvent, id: 'event-2' } },
          ],
        },
      });

      const events = await service.getEventsBatch(
        ['event-1', 'event-2'],
        mockTenantId,
        mockMicrosoftUserId
      );

      expect(events).toHaveLength(2);
      expect(mockedAxios.post).toHaveBeenCalledWith(
        'https://graph.microsoft.com/v1.0/$batch',
        expect.objectContaining({
          requests: expect.arrayContaining([
            expect.objectContaining({ method: 'GET' }),
          ]),
        }),
        expect.any(Object)
      );
    });

    it('should skip 404 responses (deleted events)', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: {
          responses: [
            { id: '0', status: 200, body: mockEvent },
            { id: '1', status: 404, body: {} },
          ],
        },
      });

      const events = await service.getEventsBatch(
        ['event-1', 'deleted-event'],
        mockTenantId,
        mockMicrosoftUserId
      );

      expect(events).toHaveLength(1);
    });

    it('should limit to 20 events per batch', async () => {
      const manyEventIds = Array.from({ length: 25 }, (_, i) => `event-${i}`);
      mockedAxios.post.mockResolvedValueOnce({
        data: {
          responses: Array.from({ length: 20 }, (_, i) => ({
            id: String(i),
            status: 200,
            body: { ...mockEvent, id: `event-${i}` },
          })),
        },
      });

      await service.getEventsBatch(manyEventIds, mockTenantId, mockMicrosoftUserId);

      const batchCall = mockedAxios.post.mock.calls[0];
      const batchPayload = batchCall[1] as { requests: Array<unknown> };
      expect(batchPayload.requests).toHaveLength(20);
    });
  });
});
