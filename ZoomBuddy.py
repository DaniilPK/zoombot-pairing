import threading
import subprocess, csv, sys

from os import system
from datetime import datetime
from time import sleep
import requests

local_filename = 'ZoomBuddy.csv'
url = 'https://docs.google.com/spreadsheets/d/1aQJ9ruOjQeThzvn4lfHGkohvvcWvd_F-/export?format=csv'


class Meeting():

    def __init__(self):
        self.file, self.csvfile = self.open_data()

        self.start()

    def start(self):
        SelectManually = input('Select Meeting manually? (Y/N) ')
        if SelectManually.lower() == 'yes' or SelectManually.lower() == 'y':
            system("cls||clear")
            self.manual()
        elif SelectManually.lower() == 'no' or SelectManually.lower() == 'n':
            print('auto start. The next check in ' + ((str)(self.calculate_initial_delay())) + ' minutes')
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
            response = requests.get(url)
            response.raise_for_status()

            with open(local_filename, 'wb') as f:
                f.write(response.content)

            try:
                file = open(local_filename, 'r')
                csvfile = csv.reader(file)
                next(csvfile)
                return file, csvfile
            except FileNotFoundError:
                print(local_filename + " Does Not Exist!")
                sleep(1)
                sys.exit(0)

        except requests.exceptions.RequestException as e:
            print(f"Ошибка при загрузке файла: {e}")
            sys.exit(0)

    # Automatically join a meeting that is +-15 minutes
    def auto(self):
        self.file, self.csvfile = self.open_data()

        day = datetime.today().weekday()
        time = int(datetime.now().strftime("%H")) * 60 + int(datetime.now().strftime("%M"))

        for row in self.csvfile:
            try:
                classtime = int(row[day + 4].split(":")[0]) * 60 + int(row[day + 4].split(":")[1])
                if (time > classtime - 10) and (time < classtime + 10):
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

    def connect(self, meetingID, passWD):
        command = "%appdata%\\Zoom\\bin\\Zoom.exe --url=zoommtg://zoom.us/join?confno=" + meetingID + "^&pwd=" + passWD
        try:
            subprocess.run(command.split(), shell=True, timeout=1)
        except subprocess.TimeoutExpired:
            pass


if __name__ == "__main__":
    Meeting()
