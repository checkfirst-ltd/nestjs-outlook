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
    description: 'The HTTP entity tag that represents the version of the object',
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
    description: 'The type of change that occurred',
    example: 'deleted',
  })
  @IsString()
  @IsNotEmpty()
  changeType: string = '';

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
  @IsNotEmpty()
  resourceData: OutlookResourceData = new OutlookResourceData();

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
