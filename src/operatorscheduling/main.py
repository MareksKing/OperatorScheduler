import argparse
import sys
import logging
from typing import Optional
import pandas as pd
from datetime import datetime, timedelta
import pytz
import os

from constants import baltic_char_map
from dotenv import dotenv_values
if os.name != 'posix':
    import win32com.client
else:
    from unittest.mock import MagicMock
    class win32:
        def __init__(self):
            self.client = client

    class client:
        def __init__(self):
            self.app = None
            pass

        @staticmethod
        def Dispatch(args):
            return app()
    
    class app:

        def __init__(self):
            pass

        @staticmethod
        def GetNamespace(args):
            return namespace_mock()

        @staticmethod
        def CreateItem(args):
            return MagicMock()


    class namespace_mock:
        def __init__(self):
            pass

        @staticmethod
        def GetDefaultFolder(args):
            return MagicMock()

    win32com = win32()


logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)
config = dotenv_values(".env")
if not config:
    config = dotenv_values("env")


class Operator:

    def __init__(self, name: str, operator_dates: list[str]):
        self.name: str = name
        self.email: str = self.create_email_from_name(name)
        self.operator_dates: list[datetime] = self.__convert_to_datetimes(operator_dates)

    def create_email_from_name(self, name: str) -> str:
        EMAIL_DOMAIN = config.get('EMAIL_DOMAIN', None)
        employee = config.get(f'EMP_{name}', None)
        if EMAIL_DOMAIN is None:
            logger.error("EMAIL_DOMAIN keyword not found in config")            
            sys.exit(1)
        if employee is None:
            logger.error(f"Employee keyword: EMP_{name} not found in config")            
            sys.exit(1)

        return f"{employee}@{EMAIL_DOMAIN}"

    def __convert_to_datetimes(self, operator_dates: list[str]) -> list[datetime]:
        # Telia date format: 08.08.2025 15:00-22:00 -> dd.mm.YYYY HH:MM-HH:MM

        converted_list: list[datetime] = []
        FORMAT = config.get("FORMAT", None)
        TIMEZONE = config.get("TIMEZONE", None)
        if not isinstance(operator_dates, list):
            operator_dates = [operator_dates]
        removed_time_range_date = [date.split("-")[0] for date in operator_dates]

        if FORMAT is None:
            logger.error("FORMAT keyword not found in config")            
            sys.exit(1)
        if TIMEZONE is None:
            logger.error("TIMEZONE keyword not found in config")            
            sys.exit(1)

        timezone = pytz.timezone(TIMEZONE)
        for date in removed_time_range_date:
            date = datetime.strptime(date, FORMAT)
            timedelta_offset = timezone.localize(date).utcoffset()
            new_date = date + timedelta_offset
            converted_list.append(new_date)

        return converted_list

    def __str__(self):
        return f"{self.name} - {self.email}"

    def __repr__(self):
        return f"{self.name} - {self.email}\n"


def get_next_operator(operator_timeline: list[tuple[str, str, datetime]], index: int):
    try:
        return operator_timeline[index]
    except IndexError:
        return None


class MeetingManager:

    def __init__(self, from_email):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
        except:
            self.outlook = None
            self.namespace = None
        
        self.store = self.get_email(from_email)
        self.location: str = "At work/Home"
        self.subject: str = "Upcomming shift"
        self.body: str = ""
        self.list_of_dates: list[datetime] = []

    def get_email(self, email):

        session = self.outlook.Session

        store = None
        for st in session.Stores:
            print(st, email)
            if st.DisplayName == email:
                store=st
                break   
        store_folder = store.GetDefaultFolder(9)
        
        return store_folder

    def make_meeting_title(self, service: pd.DataFrame, operator: tuple[str, str, datetime]) -> str:
        _, _, date = operator
        services = service.query("@date >= start and @date <= end")
        service_prefix = "/".join(services["service"].tolist())
        if not services["service"].tolist():
            return "Upcomming shift"
        return f"[{service_prefix}] Upcomming shift"

    def check_for_existing_shift(self, operator: tuple[str, str, datetime]):
        name, email, date = operator
        default_calendar = self.store.Items
        default_calendar.IncludeRecurrences = False
        start_date = date-timedelta(days=1)
        end_date = date+timedelta(days=1)
        restriction = f"[Start] >= '{start_date.strftime('%m/%d/%Y %H:%M')}' AND [End] <= '{end_date.strftime('%m/%d/%Y %H:%M')}'"
        matching_items = default_calendar.Restrict(restriction)

        return matching_items


    def create_appointment(self,
                           operator: tuple[str, str, datetime],
                           services: pd.DataFrame,
                           next_operator: tuple[str, str, datetime] | None,
                           specific_date: datetime | None = None):
        name, email, date = operator
        if next_operator is None:
            next_name, next_mail, next_date = "TBD", "TBD", "TBD"
        else:
            next_name, next_mail, next_date = next_operator
        
        if specific_date is not None:
            date = specific_date
        
        self.appointment = self.store.Items.Add()
        self.appointment.MeetingStatus = 1
        self.appointment.Subject = self.make_meeting_title(services, operator)
        self.appointment.Location = self.location
        self.appointment.Body = self.body + f"Next operator -> {next_name}"
        self.appointment.Start = date
        self.appointment.End = date + timedelta(hours=8)
        self.appointment.Recipients.Add(email)
        self.appointment.BusyStatus = 0
        self.appointment.ReminderMinutesBeforeStart = 24 * 60
        logger.info(f"Appointment created {self.appointment.Subject} - {self.appointment.Start} - {self.appointment.End} - {name}")

    def send_appointment(self):
        self.appointment.Save()
        self.appointment.Send()
        logger.info(f"Appointment {self.appointment.Subject} - {self.appointment.Start} - {self.appointment.End} sent")

    def clear_name_of_special_chars(self, text: str) -> str:
        if not text:
            return ""
        return ''.join(baltic_char_map.get(c, c) for c in text)

    def cancel_meeting(self, operator: tuple[str, str, datetime]):
        name, _, _ = operator
        name = config.get(f"EMP_{name}")
        if name is None:
            logger.error("Name was not found")            
            sys.exit(1)
        existing_meeting = self.check_for_existing_shift(operator)
        shifts = []
        for item in existing_meeting:
            if "Upcomming shift" in item.Subject:
                shifts.append(item)
        name = " ".join(name.split(".")).title()
        for item in shifts:
            recipients = self.clear_name_of_special_chars(item.RequiredAttendees)
            if name in recipients:
                logger.info(f"Found meeting: {item.Subject} {item.Start}")
                item.Delete()

    def find_meetings(self, name: str, date: Optional[datetime]):
        default_calendar = self.store.Items
        default_calendar.IncludeRecurrences = False

        if date is None:
            date = datetime.now()

        start_date = date-timedelta(days=1)
        end_date = date+timedelta(days=1)
        restriction = f"[Start] >= '{start_date.strftime('%m/%d/%Y %H:%M')}' AND [End] <= '{end_date.strftime('%m/%d/%Y %H:%M')}'"
        matching_items = default_calendar.Restrict(restriction)

        meetings = []
        shifts = []
        for item in matching_items:
            if "Upcomming shift" in item.Subject:
                shifts.append(item)

        for item in shifts:
            recipients = self.clear_name_of_special_chars(item.RequiredAttendees)
            if name in recipients:
                meetings.append(item)

        return meetings
        


    def update_meeting(self, operator_timeline: list[tuple[str, str, datetime]]):
        for operator in operator_timeline:
            name, email, date = operator
            name = config.get(f"EMP_{name}")
            if name is None:
                logger.error("Name was not found")            
                sys.exit(1)
            name = " ".join(name.split(".")).title()
            meetings = self.find_meetings(name, date)
            breakpoint()
            pass

def get_operator(name: str, agents: list) -> Operator | None:
    for agent in agents:
        if agent.name == name:
            return agent
    return None

def create_agent_list(df: pd.DataFrame, filter_date: list[str]|None = None) -> list[Operator]:

    AGENTS: list[Operator] = []
    if filter_date:
        queried_res = df.explode("covered_dates").query("covered_dates == @filter_date")
        for i, row in queried_res.iterrows():
            AGENTS.append(Operator(row.agent, row.covered_dates))
    else:
        for i, row in df.iterrows():
            AGENTS.append(Operator(row.agent, row.covered_dates))
    return AGENTS

def create_operator_timeline(agents: list[Operator]) -> list[tuple[str, str, datetime]]:

    operator_timeline: list[tuple[str, str, datetime]] = []
    for agent in agents:
        agent_timeline = [(agent.name, agent.email, date) for date in agent.operator_dates]
        operator_timeline.extend(agent_timeline)

    operator_timeline = sorted(operator_timeline, key=lambda x: x[1])
    return operator_timeline

def read_schedule(filepath: str, seperator: str = ",") -> pd.DataFrame:

    df = pd.read_csv(filepath_or_buffer=filepath, sep=seperator)
    print(df)
    date_columns = df.columns[1:]
    df['covered_dates'] = df[date_columns].apply(lambda row: row.dropna().index.tolist(), axis=1)
    filtered_df = df[['Agents/Date', 'covered_dates']]
    filtered_df = filtered_df.rename(columns={"Agents/Date": "agent"})
    return filtered_df

def find_agent_with_date(operator_timeline: list[tuple[str, str, datetime]],
                         date: datetime) -> list[tuple[str, str, datetime]]:
    """
    Finds the agents with the same operating dates as the date specified in the command line args
    Will return all tuples with the date
    """
    found_times = []
    for agent in operator_timeline:
        _, _, agent_date = agent
        if agent_date == date:
            found_times.append(agent)

    return found_times

def send_results(manager: MeetingManager,
                 operator_timeline: list[tuple[str, str, datetime]],
                 services: pd.DataFrame,
                 date: datetime | None):

    if date is None:
        for i in range(len(operator_timeline)):
            agent = operator_timeline[i]
            next_operator = get_next_operator(operator_timeline, index=i)
            manager.create_appointment(agent,services, next_operator)
            if not debug:
                manager.send_appointment()
        return

    operator_date = find_agent_with_date(operator_timeline, date)
    logger.info(f"{operator_date}")
    manager.create_appointment(operator_date[0], services, None)
    if not debug:
        manager.send_appointment()

def cancel_meeting(manager: MeetingManager,
                   operator_timeline: list[tuple[str, str, datetime]],
                   date: datetime | None):

    if date is None:
        for op_timeline in operator_timeline:
            manager.cancel_meeting(op_timeline)
        return

    operator_date = find_agent_with_date(operator_timeline, date)
    if not debug:
        manager.cancel_meeting(operator_date[0])


def main(args: argparse.Namespace, debug_):
    
    global debug
    debug = debug_
    logger.info(f"Using input file: {args.input}")
    filtered_df = read_schedule(args.input, seperator=',')

    logger.info(f"Using service timeline: {args.service}")
    service_df = pd.read_csv(args.service, sep=";")
    service_df["start"] = pd.to_datetime(service_df["start"])
    service_df["end"] = pd.to_datetime(service_df["end"])


    if args.date:
       date = [args.date]
    else:
        logger.info("No specific date specified")
        date = None

    logger.info(f"Creating agent list")
    AGENTS = create_agent_list(filtered_df, date)
    logger.info(f"Agent list created: {AGENTS}")
    manager = MeetingManager(args.email)
    
    if args.agent:
        logger.info(f"Agent specified: {args.agent}")
        operator = get_operator(args.agent, AGENTS)
        if operator is None:
            logger.error("No agent found")
            return
        operator_timeline = create_operator_timeline([operator])
        if debug:
            for operator in operator_timeline: logger.info(f"{operator}")
    else:
        logger.info("Creating operator timeline")
        operator_timeline = create_operator_timeline(AGENTS)
        logger.info("Operator timeline created")
        if debug:
            for operator in operator_timeline: logger.info(f"{operator}")

    manager.update_meeting(operator_timeline)
    SEND_BOOL = eval(args.send)
    CANCEL_BOOL = eval(args.cancel)
    if SEND_BOOL or debug:
        logger.info("Sending out the meeting reminders")
        send_results(manager, operator_timeline, service_df, date)

    if CANCEL_BOOL or debug:
        logger.info("Cancelling the meeting reminder")
        cancel_meeting(manager, operator_timeline, date)

    logger.info("Process finished")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=str, default="./agents_schedulers.csv", help="Schedule CSV file")
    parser.add_argument("--service", type=str, default="./service_timeline.csv", help="Service Main/Backup schedule")
    parser.add_argument("--agent", type=str, default=None, help="Single out one operator for scheduling")
    parser.add_argument("--date", type=str, default=None, help="Specific date to run scheduling on, format: YYYY-mm-ddTHH")
    parser.add_argument("--send", type=str, default=False, help="Send out the meeting reminders")
    parser.add_argument("--cancel", type=str, default=False, help="Cancel the meeting")
    parser.add_argument("--email", type=str, default=False, help="Email from which to send the reminder")
    args = parser.parse_args()
    main(args, debug_=False)


