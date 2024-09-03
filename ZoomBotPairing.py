import io
import json
import os
import shutil
import threading
import subprocess, csv, sys

from os import system
from datetime import datetime
from time import sleep
import requests

table_url = 'https://docs.google.com/spreadsheets/d/1aQJ9ruOjQeThzvn4lfHGkohvvcWvd_F-/export?format=csv'

VERSION = "v1.0.1"


class Meeting:

    def __init__(self):
        self.file, self.csvfile = self.open_data()

        self.weekNumber = int(input('Select Week Number: '))

        self.start()

    def start(self):
        SelectManually = input('Select Meeting manually? (Y/N) ')
        if SelectManually.lower() == 'yes' or SelectManually.lower() == 'y':
            system("cls||clear")
            self.manual()
        elif SelectManually.lower() == 'no' or SelectManually.lower() == 'n':
            print('auto start. The next check in ' + ((str)(self.calculate_initial_delay())) + ' minutes')
            self.auto()
            threading.Timer(self.calculate_initial_delay() * 60, self.schedule_auto).start()
        else:
            print('Please enter Y(yes) or N(no)')
            self.start()

    def schedule_auto(self):
        self.auto()
        print('auto. The next check in ' + ((str)(self.calculate_initial_delay())) + ' minutes')

        # Reschedule `schedule_auto` to run after the specified interval
        threading.Timer(self.calculate_initial_delay() * 60, self.schedule_auto).start()

    def calculate_initial_delay(self):
        now = datetime.now()
        current_minute = now.minute
        # Calculate delay to the next 15-minute mark
        delay = (15 - current_minute % 15)
        return delay

    def open_data(self):
        try:
            response = requests.get(table_url)
            response.raise_for_status()

            csvfile = io.StringIO(response.text)
            csv_reader = csv.reader(csvfile)
            next(csv_reader)
            return csvfile, csv_reader
        except requests.exceptions.RequestException as e:
            print(f"Ошибка при загрузке файла: {e}")
            sys.exit(0)

    # Automatically join a meeting that is +-15 minutes
    def auto(self):
        self.file, self.csvfile = self.open_data()

        day = datetime.today().weekday()
        time = int(datetime.now().strftime("%H")) * 60 + int(datetime.now().strftime("%M"))

        for row in self.csvfile:

            times = row[day + 4].split(";")
            selected_times = []

            for time_part in times:
                if time_part.startswith("I:") and self.weekNumber == 1:
                    selected_times.append(time_part[2:])
                elif time_part.startswith("II:") and self.weekNumber == 2:
                    selected_times.append(time_part[3:])
                elif not time_part.startswith("I:") and not time_part.startswith("II:"):
                    selected_times.append(time_part)

            for selected_time in selected_times:
                try:
                    classTime = int(selected_time.split(":")[0]) * 60 + int(selected_time.split(":")[1])
                    if (time > classTime - 15) and (time < classTime + 15):
                        meetingID = row[2]
                        try:
                            passWD = row[3]
                        except:
                            passWD = ""
                        self.connect(meetingID, passWD)
                except ValueError:
                    pass

    def manual(self):

        print("Choose an option below:")
        for row in self.csvfile:
            print(row[0] + " [" + row[1] + "]")
        self.file.seek(1)

        meetingname = input("Enter a Meeting: ")

        for row in self.csvfile:
            if (meetingname == row[1]):
                meetingID = row[2]
                try:
                    passWD = row[3]
                except:
                    passWD = ""
                break

        try:
            meetingID
        except NameError:
            print("Input Invalid")
            Choice = input('Try again? (Y/N) ')
            if Choice.lower() == 'yes' or Choice.lower() == 'y':
                # Yes to update
                system("cls||clear")
                self.manual()
            else:
                sys.exit(0)
            sys.exit(1)
        self.connect(meetingID, passWD)
        self.start()

    def connect(self, meetingID, passWD):
        command = "%appdata%\\Zoom\\bin\\Zoom.exe --url=zoommtg://zoom.us/join?confno=" + meetingID + "^&pwd=" + passWD
        try:
            subprocess.run(command.split(), shell=True, timeout=1)
        except subprocess.TimeoutExpired:
            pass


class UpdateFile:
    def __init__(self):
        self.URL = "https://github.com/DaniilPK/zoombot-pairing"
        self.API_URL = 'https://api.github.com/repos/DaniilPK/zoombot-pairing/releases/latest'
        self.FILE_NAME = "ZoomBotPairing.exe"

    def update_process(self):
        latest_version, download_url = self.get_latest_release()

        if int(VERSION.split("v")[1].replace('.', '')) < int(latest_version.split("v")[1].replace('.', '')):
            print(f"Найдена новая версия: {latest_version}. Обновление...")
            self.update_application(download_url)
        else:
            print("Вы используете последнюю версию.")

    def get_latest_release(self):
        response = requests.get(self.API_URL)
        release_info = response.json()
        latest_version = release_info['tag_name']
        asset = next((item for item in release_info['assets'] if item['name'] == self.FILE_NAME), None)
        if asset:
            return latest_version, asset['browser_download_url']
        return None, None

    def download_file(self, url, local_filename):
        with requests.get(url, stream=True) as r:
            with open(local_filename, 'wb') as f:
                shutil.copyfileobj(r.raw, f)

    def update_application(self, download_url):
        new_file = self.FILE_NAME + ".new"
        self.download_file(download_url, new_file)

        with open("update.bat", "w") as bat_file:
            bat_file.write(f"""
            @echo off
            timeout /t 2 /nobreak > nul
            taskkill /f /im "{self.FILE_NAME}" > nul 2>&1
            timeout /t 2 /nobreak > nul
            move /y "{new_file}" "{self.FILE_NAME}"
            start "" "{self.FILE_NAME}"
            del "%~f0"
            """)

        subprocess.Popen("update.bat", shell=True)
        sys.exit()


if __name__ == "__main__":
    UpdateFile().update_process()
    Meeting()
