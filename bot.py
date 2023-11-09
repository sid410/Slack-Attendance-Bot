import slack
import os
from pathlib import Path
from dotenv import load_dotenv
from flask import Flask, request, Response
from slackeventsapi import SlackEventAdapter
import datetime
import argparse
from attendance_handler import *
from grade_handler import *

# Parse the arguments path to excel file and number of students
parser = argparse.ArgumentParser()
parser.add_argument("xlsx_file", type=str)
parser.add_argument("--num_students", type=int, default=40)
args = parser.parse_args()

# Load the environment variables.
# Don't forget to create the .env file and update SLACK_TOKEN and SIGNING_SECRET
env_path = Path('.') / '.env'
load_dotenv(dotenv_path=env_path)

# Initialize the flask app
app = Flask(__name__)
slack_event_adapter = SlackEventAdapter(
    os.environ['SIGNING_SECRET'],'/slack/events',app)

client = slack.WebClient(token=os.environ['SLACK_TOKEN'])
BOT_ID = client.api_call("auth.test")['user_id']


# Initialize the excel file to edit
wb, ws, att = init_attendance(args.xlsx_file, str(datetime.date.today()), args.num_students)

# Format of the confirmation message to reply back
class ConfirmationMessage:
    START_TEXT = {
        'type': 'section',
        'text': {
            'type': 'mrkdwn',
            'text': (
                'Your attendance has been recorded for:'
            )
        }
    }

    DIVIDER = {'type': 'divider'}

    def __init__(self, channel, user):
        self.channel = channel
        self.user = user
        self.icon_emoji = ':robot_face:'
        self.timestamp = ''

    def get_message(self):
        return {
            'ts': self.timestamp,
            'channel': self.channel,
            'username': 'Welcome Robot!',
            'icon_emoji': self.icon_emoji,
            'blocks':[
                self.START_TEXT,
                self.DIVIDER,
                self._get_date_today()
            ]
        }
    
    def _get_date_today(self):
        checkmark = ':white_check_mark:'
        text = f'{checkmark} {str(datetime.date.today())}'
        return {'type': 'section', 'text': {'type': 'mrkdwn', 'text': text}}

# Format of the broadcast grade message
class GradeMessage:
    START_TEXT = {
        'type': 'section',
        'text': {
            'type': 'mrkdwn',
            'text': (
                'Your grade details as of'
            )
        }
    }

    DIVIDER = {'type': 'divider'}

    def __init__(self, channel, user):
        self.channel = channel
        self.user = user
        self.icon_emoji = ':robot_face:'
        self.timestamp = ''

    def message_grade_details(self, details):
        return {
            'ts': self.timestamp,
            'channel': self.channel,
            'username': 'Welcome Robot!',
            'icon_emoji': self.icon_emoji,
            'blocks':[
                self.START_TEXT,
                self._get_date_today(),
                self.DIVIDER,
                self.show_grade_details(details),
            ]
        }
    
    def show_grade_details(self, details):
        return {'type': 'section', 'text': {'type': 'mrkdwn', 
                                            'text': f"Name: {details[0]}\
                                                \nAbsent: {details[1]} times\
                                                \nParticipation: {details[2]} bonus\
                                                \nAttendance: {details[3]} %\
                                                \nLab Report: {details[5]} %\
                                                \nExam Midterm: {details[6]} %\
                                                \nMidsem Total: {details[8]} %\
                                                "}}

    def _get_date_today(self):
        text = f'{str(datetime.date.today())}'
        return {'type': 'section', 'text': {'type': 'mrkdwn', 'text': text}}

def send_confirm_message(channel, user, att):
    # Update attendance list
    att = record_attendance(att, user)

    # Send a confirmation message back to user that attendance was successfully recorded
    confirm = ConfirmationMessage(channel, user)
    message = confirm.get_message()
    response = client.chat_postMessage(**message)
    confirm.timestamp = response['ts']


# Triggered every time the /attendance command is called
@app.route('/attendance', methods=['POST'])
def message_count():
    data = request.form

    # Get the user who sent the message and if there is additional text parameter
    user_id = data.get('user_id')
    text = data.get('text')
    
    # If text parameter is save, try to save the current att list to the opened xlsx file
    # Otherwise, update attendance list and send back confirmation message
    if text == 'save':
        save_attendance(wb, ws, att, args.xlsx_file)
    else:
        send_confirm_message(f'@{user_id}', user_id, att)
    
    return Response(), 200


def send_grade_privately(channel, user, grade_list):

    # Send a confirmation message back to user that attendance was successfully recorded
    grade_msg = GradeMessage(channel, user)
    message = grade_msg.message_grade_details(grade_list)
    response = client.chat_postMessage(**message)
    grade_msg.timestamp = response['ts']

# Triggered every time the /grade command is called
@app.route('/grade', methods=['POST'])
def send_grade_info():
    data = request.form

    # Make it so only I can send the broadcast message
    user_id = data.get('user_id')
    text = data.get('text')

    if user_id == 'U057BCVHWCQ' and text == 'broadcast':
        # for loop here later
        # send_grade_privately(f'@{user_id}', user_id)
        grades_all = grades_dictionary(args.xlsx_file)
        for student in grades_all:
            print(f"sent to {student}")
            send_grade_privately(f'@{student}', student, grades_all[student])

    
    return Response(), 200

if __name__ == "__main__":
    app.run(debug=True)