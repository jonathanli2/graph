import { Injectable } from '@angular/core';
import { Client } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { AuthService } from './auth.service';
import { AlertsService } from './alerts.service';

@Injectable({
  providedIn: 'root'
})

export class GraphService {

  private graphClient: Client;
  constructor(
    private authService: AuthService,
    private alertsService: AlertsService) {

    // Initialize the Graph client
    this.graphClient = Client.init({
      authProvider: async (done) => {
        // Get the token from the auth service
        const token = await this.authService.getAccessToken()
          .catch((reason) => {
            done(reason, null);
          });

        if (token)
        {
          done(null, token);
        } else {
          done("Could not get an access token", null);
        }
      }
    });
  }

  async getSharedCalendarView(id: string, start: string, end: string, timeZone: string): Promise<MicrosoftGraph.Event[] | undefined> {
    try {
      // GET /me/calendarview?startDateTime=''&endDateTime=''
      // &$select=subject,organizer,start,end
      // &$orderby=start/dateTime
      // &$top=50
      let api = '/me/calendars/'+id+'/calendarview';
      const result =  await this.graphClient
        //.api('/me/calendarview')
        .api(api)
        .header('Prefer', `outlook.timezone="${timeZone}"`)
        .query({
          startDateTime: start,
          endDateTime: end
        })
        .select('subject,organizer,start,end')
        .orderby('start/dateTime')
        .top(50)
        .get();

      return result.value;
    } catch (error) {
      this.alertsService.addError('Could not get events', JSON.stringify(error, null, 2));
    }
    return undefined;
  }

  async getMyCalendarView(start: string, end: string, timeZone: string): Promise<MicrosoftGraph.Event[] | undefined> {
    try {
      // GET /me/calendarview?startDateTime=''&endDateTime=''
      // &$select=subject,organizer,start,end
      // &$orderby=start/dateTime
      // &$top=50
      const result =  await this.graphClient
        .api('/me/calendarview')
        .header('Prefer', `outlook.timezone="${timeZone}"`)
        .query({
          startDateTime: start,
          endDateTime: end
        })
        .select('subject,organizer,start,end')
        .orderby('start/dateTime')
        .top(50)
        .get();

      return result.value;
    } catch (error) {
      this.alertsService.addError('Could not get events', JSON.stringify(error, null, 2));
    }
    return undefined;
  }


  async getCalendars(): Promise<MicrosoftGraph.Calendar[] | undefined> {
    try {
      // GET /me/calendarview?startDateTime=''&endDateTime=''
      // &$select=subject,organizer,start,end
      // &$orderby=start/dateTime
      // &$top=50
      const result =  await this.graphClient
        .api('/me/calendars')
        .top(50)
        .get();

      return result.value;
    } catch (error) {
      this.alertsService.addError('Could not get events', JSON.stringify(error, null, 2));
    }
    return undefined;
  }

  async addEventToCalendar(newEvent: MicrosoftGraph.Event): Promise<void> {
    try {
      // POST /me/events
      await this.graphClient
        .api('/me/events')
        .post(newEvent);
    } catch (error) {
      throw Error(JSON.stringify(error, null, 2));
    }
  }

}
