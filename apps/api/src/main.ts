import { NestFactory } from "@nestjs/core";
import { json, urlencoded } from "express";
import { AppModule } from "./app.module";

async function bootstrap() {
  const app = await NestFactory.create(AppModule);
  app.use(json({ limit: "25mb" }));
  app.use(urlencoded({ extended: true, limit: "25mb" }));
  app.setGlobalPrefix("api");
  app.enableCors({ origin: ["http://localhost:5173"] });
  await app.listen(process.env.PORT ?? 3000);
}

bootstrap();
