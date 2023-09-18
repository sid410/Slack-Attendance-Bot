# Slack-Attendance-Bot
My slack bot to check students' attendance

### Start ngrok server
- ngrok http 5000
- Copy forwarding link and update slash command link in: https://api.slack.com/

### Run the virtual environment
```
py -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
```

### Prepare Excel file
Place the slack ID's of all students on the very first column under the worksheet "Classlist"

### Slash Commands
- For student to log self in attendance list, type command:
  ```/attendance```
- For teacher to save current attendance list, type command:
  ```/attendance save```
