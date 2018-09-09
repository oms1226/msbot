# -*- coding: utf-8 -*-
import os
import win32com.client as wincl

# TELEGRAM_TOKEN = '586824321:AAEvRbeyOzndEW6GuwmtNB_Pe5iUJoTDgLI'
TELEGRAM_TOKEN = ''
CHAT_ID = ''

# pip install python-telegram-bot --upgrade
def ToTelegram(str):
    global TELEGRAM_TOKEN
    global CHAT_ID

    if TELEGRAM_TOKEN == '':
        # t.me/omsStockBot
        with open('secret/telegram_token.txt', mode='r') as tokenfile:
            TELEGRAM_TOKEN = tokenfile.readline().strip()

    if TELEGRAM_TOKEN is not "":
        import telegram

        bot = telegram.Bot(token=TELEGRAM_TOKEN)

        if CHAT_ID == '':
            if os.path.exists('secret/chatid.txt'):
                with open('secret/chatid.txt', mode='r') as chatfile:
                    try:
                        CHAT_ID = int(chatfile.readline().strip())
                    except Exception as e:
                        pass

        if CHAT_ID == '':
            updates = bot.getUpdates()
            last_message = None
            for u in updates:
                if u is not None:
                    last_message = u
                # print(u.message)

            if last_message is not None:
                CHAT_ID = last_message.message.chat.id
                with open('secret/chatid.txt', mode='w') as chatfile:
                    chatfile.write("%s" % CHAT_ID)

        if CHAT_ID is not None:
            try:
                bot.sendMessage(chat_id=CHAT_ID, text=str)
            except Exception as e:
                pass


def Speak(str):
    speak = wincl.Dispatch("SAPI.SpVoice")
    speak.Speak(str)
