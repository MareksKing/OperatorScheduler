from pprint import pprint
import pandas as pd
import win32com.client
from datetime import datetime, timedelta

from dotenv import dotenv_values
from win32com.client.dynamic import CDispatch

config = dotenv_values(".env")
class Operator:

    def __init__(self, name: str, operator_dates: list[str]):
        self.name: str = name
        self.email: str = self.create_email_from_name(name)
        self.operator_dates: list[datetime] = self.__convert_to_datetimes(operator_dates)

    def create_email_from_name(self, name:str) -> str:
        EMAIL_DOMAIN = config.get('EMAIL_DOMAIN', None)
        employee = config.get(f'EMP_{name}', None)
        if EMAIL_DOMAIN is None:
            raise ValueError("EMAIL_DOMAIN keyword not found in config")
        if employee is None:
            raise ValueError(f"Employee keyword: EMP_{name} not found in config")

        return f"{employee}@{EMAIL_DOMAIN}"

    def __convert_to_datetimes(self, operator_dates: list[str]) -> list[datetime]:
        converted_list: list[datetime] = []
        #Telia date format: 08.08.2025 15:00-22:00 -> dd.mm.YYYY HH:MM-HH:MM
        removed_time_range_date = [date.split("-")[0] for date in operator_dates]
        for date in removed_time_range_date:
            converted_date:datetime = datetime.strptime(date, "%d.%m.%Y %H:%M")
            converted_list.append(converted_date)

        return converted_list

    def __str__(self):
        return f"{self.email} - {self.operator_dates}"

    def __repr__(self):
        return f"{self.email} - {self.operator_dates}"

class MeetingManager:

    def __init__(self):
        self.outlook: CDispatch = win32com.client.Dispatch("Outlook.Application")
        self.namespace: CDispatch = self.outlook.GetNamespace("MAPI")
        self.appointment: CDispatch = self.outlook.CreateItem(1)
        self.location: str = "At work/Home"
        self.subject: str = "Upcomming shift"
        self.body: str = ""

    def create_appointment(self, operator: Operator):
        self.list_of_dates: list[datetime] = operator.operator_dates
        self.appointment.Recipients.Add(operator.email)
        for date in self.list_of_dates:
            self.start_time = date
            self.end_time = date + timedelta(minutes=30)
            self.send_appointment()

    def send_appointment(self):
        self.appointment.Subject = self.subject
        self.appointment.Location = self.location
        self.appointment.Body = self.body
        self.appointment.Start = self.start_time
        self.appointment.End = self.end_time
        self.appointment.Save()
        self.appointment.Send()

df = pd.read_csv("./agents_schedulers.csv", sep=";")
date_columns = df.columns[1:]
df['covered_dates'] = df[date_columns].apply(lambda row: row.dropna().index.tolist(), axis=1)
filtered_df = df[['Agents/Date', 'covered_dates']]
filtered_df = filtered_df.rename(columns={"Agents/Date": "agent"})

AGENTS: list[Operator] = []
for i, row in filtered_df.iterrows():
    AGENTS.append(Operator(row.agent, row.covered_dates))

def get_mareks():
    for agent in AGENTS:
        if agent.name == "Mareks":
             return agent

mareks = get_mareks()
manager = MeetingManager()
manager.create_appointment(mareks)

