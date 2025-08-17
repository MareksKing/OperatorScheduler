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


def get_next_operator(operator_timeline: list, index: int):
    try:
        return operator_timeline[index]
    except IndexError:
        return None


class MeetingManager:

    def __init__(self):
        self.outlook: CDispatch = win32com.client.Dispatch("Outlook.Application")
        self.location: str = "At work/Home"
        self.subject: str = "Upcomming shift"
        self.body: str = ""
        self.list_of_dates: list[datetime] = []

    def create_appointment(self, operator: tuple[str, str, datetime], next_operator: str | None):
        name, email, date = operator
        if next_operator is None:
            n_operator = "TBD"
        else:
            n_operator = next_operator
        self.appointment: CDispatch = self.outlook.CreateItem(1)
        self.appointment.Subject = self.subject
        self.appointment.Location = self.location
        self.appointment.Body = self.body + f"Next operator -> {n_operator}"
        self.appointment.Start = date
        self.appointment.End = date + timedelta(hours=8)
        self.appointment.Recipients.Add(email)
        self.appointment.Categories = "Yellow Category"
        self.appointment.BusyStatus = 0
        self.appointment.ReminderMinutesBeforeStart = 24 * 60
        self.appointment.Save()
        self.appointment.Send()


def main():

    df = pd.read_csv(filepath_or_buffer="./agents_schedulers.csv", sep=";")
    date_columns = df.columns[1:]
    df['covered_dates'] = df[date_columns].apply(lambda row: row.dropna().index.tolist(), axis=1)
    filtered_df = df[['Agents/Date', 'covered_dates']]
    filtered_df = filtered_df.rename(columns={"Agents/Date": "agent"})

    AGENTS: list[Operator] = []
    for i, row in filtered_df.iterrows():
        AGENTS.append(Operator(row.agent, row.covered_dates))

    operator_timeline: list[tuple[str, str, datetime]] = []
    for agent in AGENTS:
        agent_timeline = [(agent.name, agent.email, date) for date in agent.operator_dates]
        operator_timeline.extend(agent_timeline)

    operator_timeline = sorted(operator_timeline, key=lambda x: x[1])

    manager = MeetingManager()
    for i in range(len(operator_timeline)):
        agent = operator_timeline[i]
        next_operator = get_next_operator(operator_timeline, index=i)
        manager.create_appointment(agent, next_operator)


if __name__ == "__main__":
    main()


