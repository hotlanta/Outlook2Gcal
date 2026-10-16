import datetime as dt
import hashlib
import json
import os
import traceback
from pathlib import Path
from zoneinfo import ZoneInfo

import emoji
import pyperclip
import win32com.client
from ctkdateentry import CTkDateEntry, CTkStringVar
from customtkinter import (
    CTk, CTkButton, CTkEntry, CTkFrame, CTkLabel, CTkScrollableFrame,
    CTkToplevel, ScalingTracker, deactivate_automatic_dpi_awareness,
    set_appearance_mode, set_default_color_theme,
)
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

APP_NAME = "Outlook2Gcal"
SCOPES = ["https://www.googleapis.com/auth/calendar"]
LOCAL_TIME_ZONE = "Europe/Budapest"
OUTLOOK_CALENDAR_FOLDER = 9

BASE_DIR = Path(__file__).resolve().parent
CONFIG_FILE = BASE_DIR / "outlook2gcal_config.json"
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"
THEME_FILE = BASE_DIR / "custom_theme.json"

LABEL_FONT = ("Trebuchet MS", 15)
BUTTON_FONT = ("Trebuchet MS", 17)
PAD = 20
ERROR_COLOR = "red"

set_appearance_mode("dark")
if THEME_FILE.exists():
    set_default_color_theme(str(THEME_FILE))
deactivate_automatic_dpi_awareness()


def load_config():
    defaults = {
        "days": 1,
        "different_day": "",
        "google_calendar_id": "primary",
        "excluded_categories": ["None"],
        "reminder_minutes": 5,
        "category_color_map": {},
    }
    if CONFIG_FILE.exists():
        try:
            with CONFIG_FILE.open("r", encoding="utf-8") as stream:
                stored = json.load(stream)
            defaults.update({key: value for key, value in stored.items() if key in defaults})
        except (OSError, json.JSONDecodeError):
            pass
    defaults["different_day"] = ""
    save_config(defaults)
    return defaults


def save_config(config):
    with CONFIG_FILE.open("w", encoding="utf-8") as stream:
        json.dump(config, stream, indent=2)


def positive_integer(value, field_name):
    result = int(value)
    if result <= 0:
        raise ValueError(f"{field_name} must be a positive whole number.")
    return result


def next_workday(day):
    candidate = day + dt.timedelta(days=1)
    while candidate.weekday() >= 5:
        candidate += dt.timedelta(days=1)
    return candidate


def selected_begin(config):
    value = config.get("different_day", "")
    if value:
        return dt.datetime.strptime(value, "%d/%m/%Y").date()
    return next_workday(dt.date.today())


def outlook_datetime(value):
    """Convert a pywin32 Outlook datetime to an aware local datetime."""
    converted = dt.datetime(
        value.year, value.month, value.day,
        value.hour, value.minute, value.second,
    )
    return converted.replace(tzinfo=ZoneInfo(LOCAL_TIME_ZONE))


def get_outlook_events(config, begin_date):
    days = positive_integer(config["days"], "Days")
    end_date = begin_date + dt.timedelta(days=days)
    excluded = {str(value).strip() for value in config.get("excluded_categories", [])}

    namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    items = namespace.GetDefaultFolder(OUTLOOK_CALENDAR_FOLDER).Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    # Outlook Restrict expects US-style dates even when Windows uses another locale.
    restriction = (
        f"[Start] >= '{begin_date.strftime('%m/%d/%Y')} 12:00 AM' AND "
        f"[Start] < '{end_date.strftime('%m/%d/%Y')} 12:00 AM'"
    )
    items = items.Restrict(restriction)

    events = []
    for item in items:
        categories = {
            part.strip()
            for part in str(getattr(item, "Categories", "")).split(",")
            if part.strip()
        }
        if excluded.intersection(categories):
            continue

        entry_id = str(getattr(item, "EntryID", "")).strip()
        subject = emoji.replace_emoji(str(getattr(item, "Subject", "") or "Untitled event"), replace="")
        is_all_day = bool(getattr(item, "AllDayEvent", False))

        if is_all_day:
            start = dt.date(item.Start.year, item.Start.month, item.Start.day)
            end = dt.date(item.End.year, item.End.month, item.End.day)
        else:
            start = outlook_datetime(item.Start)
            end = outlook_datetime(item.End)

        body = str(getattr(item, "Body", "") or "").strip()
        category_list = [
            part.strip()
            for part in str(getattr(item, "Categories", "")).split(",")
            if part.strip()
        ]
        events.append({
            "outlook_entry_id": entry_id,
            "subject": subject,
            "description": body,
            "categories": category_list,
            "location": str(getattr(item, "Location", "") or ""),
            "start": start,
            "end": end,
            "all_day": is_all_day,
        })
    return events


def get_google_credentials():
    credentials = None
    if TOKEN_FILE.exists():
        credentials = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
    elif not credentials or not credentials.valid:
        if not CREDENTIALS_FILE.exists():
            raise FileNotFoundError(
                f"Missing {CREDENTIALS_FILE.name}. Download an OAuth client JSON for a "
                "Desktop app from Google Cloud and save it beside this script with that name."
            )
        flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
        credentials = flow.run_local_server(port=0, open_browser=True)

    TOKEN_FILE.write_text(credentials.to_json(), encoding="utf-8")
    return credentials


def build_calendar_service():
    return build("calendar", "v3", credentials=get_google_credentials(), cache_discovery=False)


def outlook_sync_id(event):
    """Return a stable, API-safe identifier for one Outlook occurrence."""
    raw = f"{event['outlook_entry_id']}|{event['start'].isoformat()}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def mapped_color_id(event, category_color_map):
    """Use the first Outlook category that has a configured Google color ID."""
    for category in event.get("categories", []):
        color_id = str(category_color_map.get(category, "")).strip()
        if color_id:
            if color_id not in {str(value) for value in range(1, 12)}:
                raise ValueError(
                    f"Invalid Google color ID '{color_id}' for category '{category}'. "
                    "Use a string from 1 through 11."
                )
            return color_id
    return None


def google_event_body(event, reminder_minutes, category_color_map):
    sync_id = outlook_sync_id(event)
    body = {
        "summary": event["subject"],
        "description": event.get("description", ""),
        "location": event["location"],
        "extendedProperties": {
            "private": {
                "outlookSyncId": sync_id,
                "outlookEntryId": event["outlook_entry_id"],
            }
        },
        "reminders": {
            "useDefault": False,
            "overrides": [{"method": "popup", "minutes": reminder_minutes}],
        },
    }

    color_id = mapped_color_id(event, category_color_map)
    if color_id:
        body["colorId"] = color_id

    if event["all_day"]:
        body["start"] = {"date": event["start"].isoformat()}
        body["end"] = {"date": event["end"].isoformat()}
    else:
        body["start"] = {
            "dateTime": event["start"].isoformat(),
            "timeZone": LOCAL_TIME_ZONE,
        }
        body["end"] = {
            "dateTime": event["end"].isoformat(),
            "timeZone": LOCAL_TIME_ZONE,
        }
    return body


def find_google_event(service, calendar_id, event):
    """Find the Google event previously created for this Outlook occurrence."""
    result = service.events().list(
        calendarId=calendar_id,
        privateExtendedProperty=f"outlookSyncId={outlook_sync_id(event)}",
        maxResults=1,
        singleEvents=True,
    ).execute()
    items = result.get("items", [])
    if items:
        return items[0]

    # Compatibility with events created by the earlier script version.
    if event["outlook_entry_id"]:
        legacy_result = service.events().list(
            calendarId=calendar_id,
            privateExtendedProperty=(
                f"outlookEntryId={event['outlook_entry_id']}"
            ),
            maxResults=1,
            singleEvents=True,
        ).execute()
        legacy_items = legacy_result.get("items", [])
        if legacy_items:
            return legacy_items[0]
    return None


def sync_events(service, config, events, progress_callback=None):
    calendar_id = config.get("google_calendar_id", "primary") or "primary"
    reminder = positive_integer(
        config.get("reminder_minutes", 5), "Reminder minutes"
    )
    category_color_map = config.get("category_color_map", {})
    if not isinstance(category_color_map, dict):
        raise ValueError("category_color_map must be a JSON object.")

    created = 0
    updated = 0
    for index, event in enumerate(events, start=1):
        if progress_callback:
            progress_callback(index, len(events), event["subject"])

        event_body = google_event_body(
            event, reminder, category_color_map
        )
        existing = find_google_event(service, calendar_id, event)
        if existing:
            service.events().patch(
                calendarId=calendar_id,
                eventId=existing["id"],
                body=event_body,
                sendUpdates="none",
            ).execute()
            updated += 1
        else:
            service.events().insert(
                calendarId=calendar_id,
                body=event_body,
                sendUpdates="none",
            ).execute()
            created += 1
    return created, updated


class ModalMessage(CTkToplevel):
    def __init__(self, parent, title, message, error=False):
        super().__init__(parent)
        self.title(title)
        frame = CTkFrame(self, border_width=5)
        frame.pack(fill="both", expand=True)
        details = CTkScrollableFrame(frame, width=620, height=260 if error else 120)
        details.pack(fill="both", expand=True, padx=PAD, pady=(PAD, 10))
        CTkLabel(details, text=message, font=LABEL_FONT, wraplength=570, justify="left").pack(
            padx=10, pady=10
        )
        if error:
            CTkButton(
                frame,
                text="Copy error",
                command=lambda: pyperclip.copy(message),
                font=BUTTON_FONT,
            ).pack(side="left", padx=PAD, pady=(10, PAD))
        CTkButton(frame, text="Close", command=self.destroy, font=BUTTON_FONT).pack(
            side="right", padx=PAD, pady=(10, PAD)
        )
        center_modal(parent, self, frame, 680)


class ConfigureWindow(CTkToplevel):
    def __init__(self, parent, config):
        super().__init__(parent)
        self.title("Configure defaults")
        frame = CTkFrame(self, border_width=5)
        frame.pack(fill="both", expand=True)

        CTkLabel(frame, text="Number of days", font=LABEL_FONT).grid(
            row=0, column=0, sticky="w", padx=PAD, pady=(PAD, 5)
        )
        days = CTkEntry(frame, width=250)
        days.insert(0, str(config.get("days", 1)))
        days.grid(row=1, column=0, columnspan=2, padx=PAD, pady=(0, 10))

        CTkLabel(frame, text="Google calendar ID", font=LABEL_FONT).grid(
            row=2, column=0, sticky="w", padx=PAD, pady=(10, 5)
        )
        calendar_id = CTkEntry(frame, width=250)
        calendar_id.insert(0, str(config.get("google_calendar_id", "primary")))
        calendar_id.grid(row=3, column=0, columnspan=2, padx=PAD, pady=(0, 10))

        CTkLabel(frame, text="Reminder minutes", font=LABEL_FONT).grid(
            row=4, column=0, sticky="w", padx=PAD, pady=(10, 5)
        )
        reminder = CTkEntry(frame, width=250)
        reminder.insert(0, str(config.get("reminder_minutes", 5)))
        reminder.grid(row=5, column=0, columnspan=2, padx=PAD, pady=(0, 10))

        def save():
            try:
                config["days"] = positive_integer(days.get(), "Days")
                config["reminder_minutes"] = positive_integer(reminder.get(), "Reminder minutes")
                config["google_calendar_id"] = calendar_id.get().strip() or "primary"
                save_config(config)
                self.destroy()
            except ValueError as error:
                ModalMessage(self, "Invalid configuration", str(error), error=True)

        CTkButton(frame, text="Cancel", command=self.destroy, font=BUTTON_FONT).grid(
            row=6, column=0, padx=PAD, pady=PAD
        )
        CTkButton(frame, text="Save", command=save, font=BUTTON_FONT).grid(
            row=6, column=1, padx=PAD, pady=PAD
        )
        center_modal(parent, self, frame, 420)


class DateSelectWindow(CTkToplevel):
    def __init__(self, parent, config):
        super().__init__(parent)
        self.title("Select start date")
        frame = CTkFrame(self, border_width=5)
        frame.pack(fill="both", expand=True)
        variable = CTkStringVar(frame, value="Enter a Date")
        picker = CTkDateEntry(frame, width=150, variable=variable, font=("Roboto", 14, "bold"))
        picker.grid(row=0, column=0, columnspan=2, padx=PAD, pady=40)

        def save():
            config["different_day"] = picker.variable.get()
            save_config(config)
            self.destroy()

        CTkButton(frame, text="Cancel", command=self.destroy, font=BUTTON_FONT).grid(
            row=1, column=0, padx=PAD, pady=PAD
        )
        CTkButton(frame, text="Save", command=save, font=BUTTON_FONT).grid(
            row=1, column=1, padx=PAD, pady=PAD
        )
        center_modal(parent, self, frame, 420)


class App(CTk):
    def __init__(self, config):
        super().__init__()
        self.config = config
        self.title(APP_NAME)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        border = CTkFrame(self, border_width=5)
        border.pack(fill="both", expand=True)
        frame = CTkFrame(border)
        frame.pack(fill="both", expand=True, padx=5, pady=5)

        CTkButton(frame, text="Copy meetings", command=self.copy_meetings, font=BUTTON_FONT).grid(
            row=0, column=0, padx=PAD, pady=PAD
        )
        CTkButton(
            frame, text="Configure defaults",
            command=lambda: ConfigureWindow(self, self.config), font=BUTTON_FONT,
        ).grid(row=0, column=1, padx=PAD, pady=PAD)
        CTkButton(
            frame, text="Select a different day",
            command=lambda: DateSelectWindow(self, self.config), font=BUTTON_FONT,
        ).grid(row=0, column=2, padx=PAD, pady=PAD)
        center_window(self, border)

    def copy_meetings(self):
        try:
            begin = selected_begin(self.config)
            events = get_outlook_events(self.config, begin)
            if not events:
                ModalMessage(self, "No events", "No matching Outlook events were found.")
                return

            service = build_calendar_service()
            created, updated = sync_events(service, self.config, events)
            ModalMessage(
                self,
                "Success",
                f"Created {created} event(s) and updated {updated} event(s) in Google Calendar.",
            )
        except HttpError as error:
            ModalMessage(self, "Google Calendar API error", str(error), error=True)
        except Exception:
            ModalMessage(self, "Error", traceback.format_exc(), error=True)
        finally:
            self.config["different_day"] = ""
            save_config(self.config)


def center_modal(parent, modal, frame, min_width=0):
    center_window(modal, frame, min_width)
    modal.transient(parent)
    modal.grab_set()
    parent.wait_window(modal)


def center_window(window, frame, min_width=0, min_height=0):
    window.update_idletasks()
    frame.update_idletasks()
    try:
        scale = ScalingTracker.get_window_dpi_scaling(window)
    except Exception:
        scale = 1
    width = max(int(frame.winfo_width() * scale), min_width)
    height = max(int(frame.winfo_height() * scale), min_height)
    x = max((window.winfo_screenwidth() - width) // 2, 0)
    y = max((window.winfo_screenheight() - height) // 2, 0)
    window.geometry(f"{width}x{height}+{x}+{y}")


if __name__ == "__main__":
    application = App(load_config())
    application.mainloop()
