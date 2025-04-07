@echo off
powershell -Command "Start-Process python -ArgumentList 'noticeBot.py --chatroom \"noticebot\" --verbose' -Verb RunAs"
