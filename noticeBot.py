import time
import logging
import argparse
import requests
import pythoncom
import win32api
import win32con
import win32gui
import win32clipboard
import win32com.client
from bs4 import BeautifulSoup
from apscheduler.schedulers.background import BackgroundScheduler
from logging.handlers import TimedRotatingFileHandler

idx = 0 
# Logger setup
def set_logger():
    global botLogger
    botLogger = logging.getLogger("KakaoBot")
    botLogger.setLevel(logging.DEBUG)
    formatter = logging.Formatter("%(asctime)s | PID %(process)d | %(levelname)s | %(message)s", "%Y-%m-%d %H:%M:%S")

    handler = TimedRotatingFileHandler('./noticebot_log/webCrawling.log', when='midnight', encoding='utf-8-sig', backupCount=7)
    handler.setFormatter(formatter)
    botLogger.addHandler(handler)
    botLogger.info("Logger initialized.")

# Clipboard operations
def set_clipboard(text):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
    win32clipboard.CloseClipboard()
    time.sleep(0.2)

def send_clipboard(hwnd):
    pythoncom.CoInitialize()
    shell = win32com.client.Dispatch("WScript.Shell")
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.7)
        shell.SendKeys("^v")
        time.sleep(0.3)
        shell.SendKeys("{ENTER}")
    finally:
        pythoncom.CoUninitialize()

# Chatroom controls
def open_chatroom(chatroom_name):
    botLogger.info(f"[open_chatroom] Trying to open chatroom: {chatroom_name}")
    hwnd_kakao = win32gui.FindWindow(None, "카카오톡")
    hwnd_edit1 = win32gui.FindWindowEx(hwnd_kakao, None, "EVA_ChildWindow", None)
    hwnd_edit2_1 = win32gui.FindWindowEx(hwnd_edit1, None, "EVA_Window", None)
    hwnd_edit2_2 = win32gui.FindWindowEx(hwnd_edit1, hwnd_edit2_1, "EVA_Window", None)
    hwnd_edit3 = win32gui.FindWindowEx(hwnd_edit2_2, None, "Edit", None)

    if hwnd_edit3 == 0:
        botLogger.error("[open_chatroom] Failed to find chatroom search box.")
        return False

    win32api.SendMessage(hwnd_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(1)
    send_key(hwnd_edit3, "{ENTER}")
    time.sleep(1)
    botLogger.info(f"[open_chatroom] Chatroom '{chatroom_name}' opened.")
    return True

def clean_chatroom():
    hwnd_kakao = win32gui.FindWindow(None, "카카오톡")
    hwnd_edit1 = win32gui.FindWindowEx(hwnd_kakao, None, "EVA_ChildWindow", None)
    hwnd_edit2_1 = win32gui.FindWindowEx(hwnd_edit1, None, "EVA_Window", None)
    hwnd_edit2_2 = win32gui.FindWindowEx(hwnd_edit1, hwnd_edit2_1, "EVA_Window", None)
    hwnd_edit3 = win32gui.FindWindowEx(hwnd_edit2_2, None, "Edit", None)
    win32api.SendMessage(hwnd_edit3, win32con.WM_SETTEXT, 0, '')
    return hwnd_edit3

def send_key(hwnd, key):
    pythoncom.CoInitialize()
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(1)
        shell.SendKeys(key)
    finally:
        pythoncom.CoUninitialize()

# Send messages to KakaoTalk
def kakao_sendtext(chatroom_name, noticeLists):
    botLogger.info(f"[kakao_sendtext] Sending {len(noticeLists)} messages to '{chatroom_name}'")
    hwndMain = win32gui.FindWindow(None, chatroom_name)
    if hwndMain == 0:
        botLogger.error(f"[kakao_sendtext] Cannot find chatroom window '{chatroom_name}'")
        return

    hwndEdit = win32gui.FindWindowEx(hwndMain, None, "RICHEDIT50W", None)
    if hwndEdit == 0:
        botLogger.error(f"[kakao_sendtext] Failed to find chat input box for '{chatroom_name}'")
        return

    win32gui.SetForegroundWindow(hwndMain)
    time.sleep(1)

    for notice in noticeLists:
        message = f"📢 [공지사항] {notice['date']}\n🔹 제목: {notice['title']}\n🔗 링크: {notice['link']}"
        set_clipboard(message)
        send_clipboard(hwndMain)
        botLogger.info(f"[kakao_sendtext] Message sent: {message}")
        time.sleep(1.5)

    botLogger.info(f"[kakao_sendtext] Completed sending messages to '{chatroom_name}'")
    send_key(hwndMain, "{ESC}")
    botLogger.info(f"[kakao_close_chatroom] Closed the chatroom '{chatroom_name}'")

# Crawl DWU notice

def get_dwu_notice():
    global idx
    url = 'https://www.dongduk.ac.kr/www/contents/kor-noti.do?gotoMenuNo=kor-noti'
    base_url = 'https://www.dongduk.ac.kr/www/contents/kor-noti.do?schM=view&page=1&viewCount=10&id='

    response = requests.get(url)
    if response.status_code != 200:
        botLogger.error(f"[get_dwu_notice] Failed to fetch notices. HTTP {response.status_code}")
        return []

    soup = BeautifulSoup(response.text, 'html.parser')
    elements = soup.select('ul.board-basic li > dl')

    notice_set = []
    existing_ids = set()

    for element in elements:
        notice_id = int(element.a.get('onclick').split("'")[1])
        if notice_id in existing_ids:
            continue
        existing_ids.add(notice_id)

        title = element.a.text.strip()
        date = element.find_all('span', 'p_hide')[1].text
        notice_set.append({"id": notice_id, "title": title, "date": date, "link": f"{base_url}{notice_id}"})

    new_notices = [n for n in notice_set if n["id"] > idx]
    if new_notices:
        new_notices.sort(key=lambda x: x["id"])
        idx = new_notices[-1]["id"]
        botLogger.info(f"[get_dwu_notice] {len(new_notices)} new notices fetched.")
        return new_notices

    botLogger.info("[get_dwu_notice] No new notices found.")
    return []

# Job to run periodically
def job(chatroom_name):
    botLogger.info("[job] Running scheduled job...")
    noticeList = get_dwu_notice()
    if not noticeList:
        botLogger.info(f"[job] No new notices for '{chatroom_name}'")
        return

    if open_chatroom(chatroom_name):
        kakao_sendtext(chatroom_name, noticeList)
        if clean_chatroom() == 0:
            botLogger.error("[job] Failed to clean chatroom.")

    botLogger.info("[job] Job completed.")

# Main
def main():
    parser = argparse.ArgumentParser(description='Notice Bot for Dongduk Women\'s University')
    parser.add_argument('--chatroom', type=str, required=True, help='Chatroom name')
    parser.add_argument('--verbose', action='store_true', help='Verbose output')
    args = parser.parse_args()

    if args.verbose:
        print(f"Chatroom name: {args.chatroom}")

    set_logger()
    botLogger.info("Bot is starting...")

    sched = BackgroundScheduler()
    sched.start()
    sched.add_job(job, 'interval', minutes=15, args=[args.chatroom])

    while True:
        botLogger.debug("[main] Bot is running...")
        time.sleep(900)

if __name__ == '__main__':
    main()