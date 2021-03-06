import { Component, OnInit } from '@angular/core';

import { AuthService } from '../auth.service';
import { GraphService } from '../graph.service';
import { AlertsService } from '../alerts.service';
import { NewEvent } from './new-event';

@Component({
  selector: 'app-new-event',
  templateUrl: './new-event.component.html',
  styleUrls: ['./new-event.component.css']
})
export class NewEventComponent implements OnInit {

  model = new NewEvent();

  constructor(
    private authService: AuthService,
    private graphService: GraphService,
    private alertsService: AlertsService) { }

  ngOnInit(): void {
  }

  onSubmit(): void {
    const timeZone = this.authService.user?.timeZone ?? 'UTC';
    const graphEvent = this.model.getGraphEvent(timeZone);

    this.graphService.addEventToCalendar(graphEvent)
      .then(() => {
        this.alertsService.addSuccess('Event created.');
      }).catch(error => {
        this.alertsService.addError('Error creating event.', error.message);
      });
  }
}
