import { config } from 'dotenv';
import * as path from 'path';

// Load environment variables
const result = config({
  path: path.resolve(process.cwd(), '.env'),
});

if (result.error) {
  console.error('Error loading .env file:', result.error);
  process.exit(1);
}

console.log('Environment variables loaded successfully!');
console.log('Checking required variables:');
console.log({
  MICROSOFT_CLIENT_ID: process.env.MICROSOFT_CLIENT_ID ? '✅ Set' : '❌ Not set',
  MICROSOFT_CLIENT_SECRET: process.env.MICROSOFT_CLIENT_SECRET ? '✅ Set' : '❌ Not set',
  BACKEND_BASE_URL: process.env.BACKEND_BASE_URL ? '✅ Set' : '❌ Not set',
  MICROSOFT_REDIRECT_PATH: process.env.MICROSOFT_REDIRECT_PATH ? '✅ Set' : '❌ Using default',
  MICROSOFT_BASE_PATH: process.env.MICROSOFT_BASE_PATH ? '✅ Set' : '❌ Using default',
});

// Check if required variables are missing
const missingRequired = [
  'MICROSOFT_CLIENT_ID',
  'MICROSOFT_CLIENT_SECRET',
  'BACKEND_BASE_URL',
].filter(key => !process.env[key]);

if (missingRequired.length > 0) {
  console.error('\n❌ Missing required environment variables:', missingRequired.join(', '));
  console.error('Please check your .env file and make sure all required variables are set.');
  process.exit(1);
}

console.log('\n✅ All required environment variables are set!'); 