import { NestFactory } from '@nestjs/core';
import { ValidationPipe } from '@nestjs/common';
import { NestExpressApplication } from '@nestjs/platform-express';
import { join } from 'path';
import { AppModule } from './app.module';

async function bootstrap() {
  const app = await NestFactory.create<NestExpressApplication>(AppModule);

  // Serve static files from the public directory
  app.useStaticAssets(join(__dirname, '..', 'public'));

  // Apply global validation
  app.useGlobalPipes(new ValidationPipe({
    whitelist: true,
    transform: true,
  }));

  await app.listen(3000);
  console.log(`Application is running on: ${await app.getUrl()}`);
  console.log(`Demo page available at: ${await app.getUrl()}/index.html`);
}

// Handle bootstrap errors
bootstrap().catch((error: unknown) => {
  console.error('Failed to start application:', error);
  process.exit(1);
}); 