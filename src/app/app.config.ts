// src/app/app.config.ts
import { ApplicationConfig, provideZoneChangeDetection } from '@angular/core';
import { provideRouter } from '@angular/router';
import { provideClientHydration, withEventReplay } from '@angular/platform-browser';
import { routes } from './app.routes';

// AngularFire imports - IMPORTANT: For Angular 19, use the new modular API
import { initializeApp, provideFirebaseApp } from '@angular/fire/app';
import { getAuth, provideAuth } from '@angular/fire/auth';
import { getFirestore, provideFirestore } from '@angular/fire/firestore';
import { environment } from '../environments/environment';

export const appConfig: ApplicationConfig = {
  providers: [
    // Zone.js for change detection (Angular 19)
    provideZoneChangeDetection({ eventCoalescing: true }),
    
    // Router
    provideRouter(routes),
    
    // Client hydration (for SSR)
    provideClientHydration(withEventReplay()),
    
    // Firebase setup using the new modular API for Angular 19
    provideFirebaseApp(() => initializeApp(environment.firebase)),
    provideAuth(() => getAuth()),
    provideFirestore(() => getFirestore()),
  ]
};