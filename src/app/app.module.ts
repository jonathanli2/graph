import {
  IPublicClientApplication,
  PublicClientApplication,
  BrowserCacheLocation,
} from '@azure/msal-browser';
import { MsalModule, MsalService, MSAL_INSTANCE } from '@azure/msal-angular';
import { OAuthSettings } from '../oauth';
import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { FormsModule } from '@angular/forms';
import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { NgbModule } from '@ng-bootstrap/ng-bootstrap';
import { NavBarComponent } from './nav-bar/nav-bar.component';
import { HomeComponent } from './home/home.component';
import { AlertsComponent } from './alerts/alerts.component';
import { CalendarComponent } from './calendar/calendar.component';

let msalInstance: IPublicClientApplication | undefined = undefined;

export function MSALInstanceFactory(): IPublicClientApplication {
  msalInstance =
    msalInstance ??
    new PublicClientApplication({
      auth: {
        clientId: OAuthSettings.appId,
        redirectUri: OAuthSettings.redirectUri,
        postLogoutRedirectUri: OAuthSettings.redirectUri,
      },
      cache: {
        cacheLocation: BrowserCacheLocation.LocalStorage,
      },
    });

  return msalInstance;
}

@NgModule({
  declarations: [AppComponent, NavBarComponent, HomeComponent, AlertsComponent, CalendarComponent],
  imports: [
    BrowserModule,
    FormsModule,
    AppRoutingModule,
    NgbModule,
    MsalModule,
  ],
  providers: [
    {
      provide: MSAL_INSTANCE,
      useFactory: MSALInstanceFactory,
    },
    MsalService,
  ],
  bootstrap: [AppComponent],
})
export class AppModule {}
