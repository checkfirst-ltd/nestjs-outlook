import { IsNotEmpty, IsString, IsOptional, IsObject, IsArray } from 'class-validator';
import { ApiProperty } from '@nestjs/swagger';

/**
 * DTO representing the resource data for Outlook webhook notifications
 */
export class OutlookResourceData {
  @ApiProperty({
    description: 'The OData entity type in Microsoft Graph',
    example: '#microsoft.graph.event',
    required: false,
  })
  @IsString()
  @IsOptional()
  '@odata.type'?: string;

  @ApiProperty({
    description: 'The OData identifier of the object',
    example: 'https://graph.microsoft.com/v1.0/users/{userId}/events/{eventId}',
    required: false,
  })
  @IsString()
  @IsOptional()
  '@odata.id'?: string;

  @ApiProperty({
    description: 'The OData entity tag that represents the version of the object',
    example: 'W/"ZWRafd0rFkORSLqrpwPMEQlFkSo="',
    required: false,
  })
  @IsString()
  @IsOptional()
  '@odata.etag'?: string;

  @ApiProperty({
    description: 'The identifier of the object',
    example: 'AAMkADI5MAAIT3drCAAA=',
  })
  @IsString()
  @IsNotEmpty()
  id: string = '';

  @ApiProperty({
    description: 'The user ID associated with the resource',
    example: 123,
  })
  userId?: number;

  @ApiProperty({
    description: 'The subscription ID that triggered this notification',
    example: '08ee466c-5ceb-4af2-a98f-aea3316a854c',
  })
  subscriptionId?: string;

  @ApiProperty({
    description: 'The resource path that changed',
    example: '/me/messages/AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAAAYbuK-AAA=',
  })
  resource?: string;

  @ApiProperty({
    description: 'The type of change that occurred',
    example: 'created',
  })
  changeType?: string;

  @ApiProperty({
    description: 'Additional data for the resource (like email content for new emails)',
    required: false,
    type: 'object',
  })
  @IsObject()
  @IsOptional()
  data?: Record<string, unknown>;

  // We don't add API decorators for the index signature
  [key: string]: unknown;
}

/**
 * DTO representing a notification item from Microsoft Graph webhook
 */
export class OutlookWebhookNotificationItemDto {
  @ApiProperty({
    description: 'The ID of the webhook subscription',
    example: '08ee466c-5ceb-4af2-a98f-aea3316a854c',
  })
  @IsString()
  @IsNotEmpty()
  subscriptionId: string = '';

  @ApiProperty({
    description: 'The date and time when the subscription expires',
    example: '2019-09-16T02:17:10Z',
  })
  @IsString()
  @IsNotEmpty()
  subscriptionExpirationDateTime: string = '';

  @ApiProperty({
    description: 'The type of change that occurred (created, updated, deleted). May be absent for lifecycle events.',
    example: 'deleted',
    required: false,
  })
  @IsString()
  @IsOptional()
  changeType?: string;

  @ApiProperty({
    description: 'Lifecycle event type (missed, subscriptionRemoved, reauthorizationRequired)',
    example: 'reauthorizationRequired',
    required: false,
  })
  @IsString()
  @IsOptional()
  lifecycleEvent?: "missed" | "subscriptionRemoved" | "reauthorizationRequired";

  @ApiProperty({
    description: 'The type of resource that changed',
    example: 'event',
  })
  @IsString()
  @IsNotEmpty()
  resource: string = '';

  @ApiProperty({
    description: 'The data of the resource that changed',
    type: OutlookResourceData,
  })
  @IsObject()
  @IsOptional()
  resourceData?: OutlookResourceData;

  @ApiProperty({
    description: 'The unique identifier for the client state',
    example: 'c75831bd-fad3-4191-9a66-280a48528679',
    required: false,
  })
  @IsString()
  @IsOptional()
  clientState?: string;

  @ApiProperty({
    description: 'The tenant ID',
    example: 'bb8775a4-4d8c-4b2d-9411-b77d0123456',
    required: false,
  })
  @IsString()
  @IsOptional()
  tenantId?: string;
}

/**
 * DTO representing a notification from Microsoft Graph webhook
 */
export class OutlookWebhookNotificationDto {
  @ApiProperty({
    description: 'Array of notification items',
    type: [OutlookWebhookNotificationItemDto],
  })
  @IsArray()
  value: OutlookWebhookNotificationItemDto[] = [];
}
