# StudyTime

## Goal
Originally designed to help me focus on studying more, this script enables me to do more work. It taps into Outlook directly without needing credentials. It finds a free 60 minutes or 30 minutes timebox and creates a calendar invite on your primary calendar. This should help with making sure you have adequate time to study while also making sure people don't overbook your designated time. Ideally, this script should be run on Monday morning and it will hunt your calendar for free time.

## Config
Using the json file, you can specify your working hours. Based on that, the script will see what's open.

You can specify if you want to setup a 60 minute or 30 minute meeting. The script will automatically resize meetings from 60 minutes to 30 minutes if it finds no open 60 minute spots. If no 30 minute spots, it will give up for that day.

## Feature Backlog
- Prioritize a timebox NOT next to another timebox. We don't want to jump from one meeting into direct studying unless we have too. Prioritize times where I have padding between meetings.


### Other Considerations
I had a very hard time figuring out recurring meetings. This guy summed up how to visualize the recurring meetings and normal one-off meetings with this solutions: https://stackoverflow.com/a/12603773
