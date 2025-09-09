import os.path
import datetime as dt

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/calendar"]

class GoogleCalendarManager:
    def __init__(self):
        self.service = self._authenticate()

    def _authenticate(self):
        creds = None
        creds_file = "client_secret_131470428704-mboeav2papfa0c1cbrv2b9cklqgg7arh.apps.googleusercontent.com.json"

        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(creds_file, SCOPES)
                creds = flow.run_local_server(port=0)

            # Save the credentials for the next run
            with open("token.json", "w") as token:
                token.write(creds.to_json())

        return build("calendar", "v3", credentials=creds)

    def _get_recurrence_rule(self, frequency):
        """Converts human-readable frequency to Google Calendar API RRULE."""
        rules = {
            "Diario": "RRULE:FREQ=DAILY;BYDAY=MO,TU,WE,TH,FR",
            "Lunes, Miércoles y Viernes": "RRULE:FREQ=WEEKLY;BYDAY=MO,WE,FR",
            "Martes y Jueves": "RRULE:FREQ=WEEKLY;BYDAY=TU,TH",
            "Semanal (Viernes)": "RRULE:FREQ=WEEKLY;BYDAY=FR",
            # Add other frequencies here if needed
        }
        return [rules.get(frequency, "")]

    def list_upcoming_events(self, max_results=10):
        now = dt.datetime.utcnow().isoformat() + "Z"
        tomorrow = (dt.datetime.now() + dt.timedelta(days=5)).replace(hour=23, minute=59, second=0, microsecond=0).isoformat() + "Z"

        try:
            events_result = self.service.events().list(
                calendarId='primary', timeMin=now, timeMax=tomorrow,
                maxResults=max_results, singleEvents=True,
                orderBy='startTime'
            ).execute()
            events = events_result.get('items', [])

            if not events:
                print('No se encontraron eventos próximos.')
            else:
                for event in events:
                    start = event['start'].get('dateTime', event['start'].get('date'))
                    print(start, event['summary'], event['id'])
            return events
        except HttpError as error:
            print(f"Ha ocurrido un error al listar eventos: {error}")
            return []

    def create_event(self, summary, start_time, end_time, timezone, attendees=None, frequency=None):
        event = {
            'summary': summary,
            'start': {
                'dateTime': start_time,
                'timeZone': timezone,
            },
            'end': {
                'dateTime': end_time,
                'timeZone': timezone,
            },
            "reminders": {
                "useDefault": True
            }
        }
        
        # Add recurrence rule if frequency is provided
        if frequency:
            recurrence_rule = self._get_recurrence_rule(frequency)
            if recurrence_rule:
                event['recurrence'] = recurrence_rule
        
        # Add attendees and send email invitations
        if attendees:
            event["attendees"] = [{"email": email} for email in attendees]
            event['sendUpdates'] = 'all'

        try:
            event = self.service.events().insert(
                calendarId="primary", 
                body=event,
                sendUpdates='all'
            ).execute()
            print(f"Evento creado: {event.get('htmlLink')}")
            return event
        except HttpError as error:
            print(f"Ha ocurrido un error: {error}")
            return None

    def update_event(self, event_id, summary=None, start_time=None, end_time=None, attendees=None):
        try:
            event = self.service.events().get(calendarId='primary', eventId=event_id).execute()

            if summary:
                event['summary'] = summary
            if start_time:
                event['start']['dateTime'] = start_time.strftime('%Y-%m-%dT%H:%M:%S')
            if end_time:
                event['end']['dateTime'] = end_time.strftime('%Y-%m-%dT%H:%M:%S')
            
            # Update attendees and send invitations
            if attendees:
                event["attendees"] = [{"email": email} for email in attendees]
                event['sendUpdates'] = 'all'

            updated_event = self.service.events().update(
                calendarId='primary', 
                eventId=event_id, 
                body=event,
                sendUpdates='all'
            ).execute()
            return updated_event
        except HttpError as error:
            print(f"Ha ocurrido un error al actualizar evento: {error}")
            return None

    def delete_event(self, event_id):
        try:
            self.service.events().delete(calendarId='primary', eventId=event_id).execute()
            return True
        except HttpError as error:
            print(f"Ha ocurrido un error al eliminar evento: {error}")
            return False

# Example Usage
# calendar_manager = GoogleCalendarManager()
# new_task_data = {
#     "title": "Reunión de Equipo de Marketing",
#     "assignedTo": "correo@ejemplo.com",
#     "team": "Marketing",
#     "frequency": "Diario (L-V)",
#     "description": "Revisión diaria del estado de los proyectos."
# }
# calendar_manager.create_event_from_task(new_task_data)
