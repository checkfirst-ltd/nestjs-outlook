import { NestFactory } from '@nestjs/core';
import { ValidationPipe } from '@nestjs/common';
import { NestExpressApplication } from '@nestjs/platform-express';
import { join } from 'path';
import { AppModule } from './app.module';
import { startZrokTunnel } from './zrok-autostart';

async function bootstrap() {
  const app = await NestFactory.create<NestExpressApplication>(AppModule);

  // Serve static files from the public directory
  app.useStaticAssets(join(__dirname, '..', 'public'));

  // Apply global validation
  app.useGlobalPipes(new ValidationPipe({
    whitelist: true,
    transform: true,
  }));

  const port = process.env.PORT ?? 8888;
  await app.listen(port);
  console.log(`Application is running on: ${await app.getUrl()}`);
  console.log(`Demo page available at: ${await app.getUrl()}/index.html`);

  // Opt-in (ZROK_AUTOSTART=true): bridge a fixed public HTTPS URL to this port so
  // Microsoft can reach the OAuth callback + webhooks — no second terminal. See ZROK.md.
  startZrokTunnel(port);
}

// Handle bootstrap errors
bootstrap().catch((error: unknown) => {
  console.error('Failed to start application:', error);
  process.exit(1);
}); 