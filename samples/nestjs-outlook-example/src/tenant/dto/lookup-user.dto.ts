import { ApiProperty } from '@nestjs/swagger';
import { IsEmail, IsNotEmpty, IsString } from 'class-validator';

export class LookupUserDto {
  @ApiProperty({
    description: 'Email address to look up in the tenant',
    example: 'john.doe@contoso.com',
  })
  @IsEmail()
  @IsNotEmpty()
  email!: string;
}

export class RegisterUserMappingDto {
  @ApiProperty({
    description: 'External user ID from your application',
    example: 'user-123',
  })
  @IsString()
  @IsNotEmpty()
  externalUserId!: string;

  @ApiProperty({
    description: 'Email address of the Microsoft user to map',
    example: 'john.doe@contoso.com',
  })
  @IsEmail()
  @IsNotEmpty()
  email!: string;
}
