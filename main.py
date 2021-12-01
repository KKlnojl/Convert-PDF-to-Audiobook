import pyttsx3
import pandas as pd
import openpyxl  # 需有此才能順利開啟xlsx
import tkinter as tk
import PyPDF2
# 也可用pdfplumber
import os
import boto3  # AWS SDK for Python (Boto3)


# ---------------------------------- 常數設定 ---------------------------------- #
KEY_WORDS = "Many online quizzes at URL below\n "
KEY_END_WORDS = "\n Sources"

# AWS
AWS_KEY_ID = os.environ.get("AWS_ID")
AWS_ACCESS_KEY = os.environ.get("AWS_ACCESS_KEY")


# ---------------------------------- Class ---------------------------------- #
# Create a Reading Breaking News class
class ReadBreakingNews(PyPDF2.PdfFileReader):
    def __init__(self, f):
        super().__init__(stream=f, strict=True, warndest=None, overwriteWarnings=True)
        self.origin = self.read_all()
        self.title, self.published_date, self.content = self.find_paragraph(start=KEY_WORDS, end=KEY_END_WORDS)
        self.news = self.published_date + "\n" + self.title + "\n" + self.content

    def read_all(self):
        text = []
        for page in range(self.numPages):
            text.append(self.getPage(page).extractText())
        return text

    def find_paragraph(self, start=None, end=None):
        """
        :param start: Put the key words which is put before title.
        :param end: ut the key words which is put after content.
        Usually, it won't be changed unless the website change the basic form.
        :return: title, date, content, sentences
        """
        # 依排版確認pdf關鍵字後，找到其index(title_start_from)，再找第一次出現'\n '(title_end_from)即後面會接日期的地方
        # 取出此區間的字串，即為title
        # 再以此方式，陸續找出date 和content
        title_start_from = self.origin[0].find(start)  # return index
        title_start_from = title_start_from + len(start)
        title_end_from = self.origin[0][title_start_from:].find("\n ")
        title_end_from = title_start_from + title_end_from
        title = self.origin[0][title_start_from:title_end_from]
        title = title.replace("\n", "")
        date_start_from = self.origin[0][title_end_from:].find("\n ")
        date_start_from = title_end_from + date_start_from
        # +8 係針對'\n '加上'\n, 2021'
        date_end_from = self.origin[0][title_end_from:].find("\n, ") + 8
        date_end_from = date_start_from + date_end_from
        published_date = self.origin[0][date_start_from:date_end_from]
        published_date = published_date.replace("\n", "").lstrip()
        content_end_from = self.origin[0].find(end)
        content = self.origin[0][date_end_from:content_end_from].lstrip()
        content = content.replace("\n", "")
        content = content.replace('\"', '')
        return title, published_date, content


# ---------------------------------- UI setting -Main Window ---------------------------------- #


# ---------------------------------- Preliminary - Tackle PDF ---------------------------------- #
# Read pdf by PyPDF2
file = open("example.pdf", "rb")
pdf = ReadBreakingNews(file)
news = pdf.news
print(news)  # 留著，可以確認完整性，以確認版本是否有變動

# ---------------------------------- Convert PDF to Audio ---------------------------------- #
# Method 1: Offline - 轉換pdf內容成audio
# engine = pyttsx3.init()
# voices = engine.getProperty("voices")  # 取得發音原
# rate = engine.getProperty("rate")
# engine.setProperty("voice", voices[2].id)  # 設定發音員
# engine.setProperty("rate", 130)  # 設定語速，預設為每分鐘200
# engine.say(news)  # 將轉換成音檔的文章放入
# engine.save_to_file(news, "python_example_audio.mp3")  # 儲存音檔
# engine.runAndWait()  # 執行(要放最後)

# Method 2: Online - Use AWS and boto3 to create the audio
# https://boto3.amazonaws.com/v1/documentation/api/latest/reference/services/polly.html?highlight=polly#Polly.Client.synthesize_speech
polly_client = boto3.Session(
    aws_access_key_id=AWS_KEY_ID,
    aws_secret_access_key=AWS_ACCESS_KEY,
    region_name='us-west-2').client('polly')  # polly為AWS提供的 text-to-speech(TTS)服務

response = polly_client.synthesize_speech(VoiceId='Joanna',
                                          OutputFormat='mp3',
                                          LanguageCode='en-US',
                                          Text=news,
                                          Engine='neural')

file = open('polly_example_audio.mp3', 'wb')  # 建立一個mp3檔
file.write(response['AudioStream'].read())
file.close()
