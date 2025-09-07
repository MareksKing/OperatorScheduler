import pandas as pd
import win32com.client
from datetime import datetime, timedelta
import pytz

from dotenv import dotenv_values
from win32com.client.dynamic import CDispatch


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
            raise ValueError("EMAIL_DOMAIN keyword not found in config")
        if employee is None:
            raise ValueError(f"Employee keyword: EMP_{name} not found in config")

        return f"{employee}@{EMAIL_DOMAIN}"

    def __convert_to_datetimes(self, operator_dates: list[str]) -> list[datetime]:
        # Telia date format: 08.08.2025 15:00-22:00 -> dd.mm.YYYY HH:MM-HH:MM

        converted_list: list[datetime] = []
        FORMAT = config.get("FORMAT", None)
        TIMEZONE = config.get("TIMEZONE", None)
        removed_time_range_date = [date.split("-")[0] for date in operator_dates]

        if FORMAT is None:
            raise ValueError("FORMAT keyword not found in config")
        if TIMEZONE is None:
            raise ValueError("TIMEZONE keyword not found in config")

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
        return f"{self.name} - {self.email}"


def get_next_operator(operator_timeline: list[tuple[str, str, datetime]], index: int):
    try:
        return operator_timeline[index]
    except IndexError:
        return None


class MeetingManager:

    def __init__(self):
        self.outlook: CDispatch = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.location: str = "At work/Home"
        self.subject: str = "Upcomming shift"
        self.body: str = ""
        self.list_of_dates: list[datetime] = []

    def check_for_existing_shift(self, operator: tuple[str, str, datetime]):
        name, email, date = operator
        default_calendar = self.namespace.GetDefaultFolder(9).Items
        default_calendar.IncludeRecurrences = False
        start_date = date
        end_date = date+timedelta(days=1)

        restriction = f"[Start] >= '{start_date.strftime('%m/%d/%Y %H:%M')}' AND [Subject] = {self.subject} AND [End] <= '{end_date.strftime('%m/%d/%Y %H:%M')}'"
        matching_items = default_calendar.Restrict(restriction)
        if matching_items.Recipients.name != name:
            matching_items.CancelMeeting()
            matching_items.Delete()

        
        for item in matching_items:
            recipients = [rec.name for rec in item.Recipients]
            print(f"Deleting: {item.Subject} at {item.Start} - attendee: {recipients}")

        return matching_items


    def create_appointment(self, operator: tuple[str, str, datetime], next_operator: tuple[str, str, datetime] | None, specific_date: datetime | None = None):
        name, email, date = operator
        if next_operator is None:
            next_name, next_mail, next_date = "TBD", "TBD", "TBD"
        else:
            next_name, next_mail, next_date = next_operator
        
        if specific_date is not None:
            date = specific_date
        self.appointment: CDispatch = self.outlook.CreateItem(1)
        self.appointment.MeetingStatus = 1
        self.appointment.Subject = self.subject
        self.appointment.Location = self.location
        self.appointment.Body = self.body + f"Next operator -> {next_name}"
        self.appointment.Start = date
        self.appointment.End = date + timedelta(hours=8)
        self.appointment.Recipients.Add(email)
        self.appointment.BusyStatus = 0
        self.appointment.ReminderMinutesBeforeStart = 24 * 60
        existing_meeting = self.check_for_existing_shift(operator)
#        self.appointment.Save()
#        self.appointment.Send()

def get_operator(name: str, agents: list) -> Operator | None:
    for agent in agents:
        if agent.name == name:
            return agent
    return None

def create_agent_list(df: pd.DataFrame) -> list[Operator]:

    AGENTS: list[Operator] = []
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

def main():

    df = pd.read_csv(filepath_or_buffer="./agents_schedulers.csv", sep=",")
    print(df)
    date_columns = df.columns[1:]
    df['covered_dates'] = df[date_columns].apply(lambda row: row.dropna().index.tolist(), axis=1)
    filtered_df = df[['Agents/Date', 'covered_dates']]
    filtered_df = filtered_df.rename(columns={"Agents/Date": "agent"})

    AGENTS = create_agent_list(filtered_df)

    operator_timeline = create_operator_timeline(AGENTS)
    
    manager = MeetingManager()
    mareks = get_operator("Mareks", AGENTS)
    if mareks is None:
        print("No agent found")
        return
    skaidrite = get_operator("Skaidrite", AGENTS)

    tuples = (mareks.name, mareks.email, datetime(2025, 7, 24, 10, 0))
    manager.create_appointment(tuples, tuples)
   # for i in range(len(operator_timeline)):
   #     agent = operator_timeline[i]
   #     next_operator = get_next_operator(operator_timeline, index=i)
   #     manager.create_appointment(agent, next_operator)


if __name__ == "__main__":
    main()


