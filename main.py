print("[#] Attendance Distribution BOT ")
print("[#] author : https://github.com/akhilrajs")
print("")
print("")
print("[#] importing modules ...")


try:
	from selenium import webdriver
	from selenium.webdriver.chrome.options import Options
	from selenium.webdriver.support.ui import WebDriverWait
	from selenium.webdriver.support import expected_conditions as EC
	from selenium.webdriver.common.by import By
	from webdriver_manager.chrome import ChromeDriverManager
	from time import sleep
	from urllib.parse import quote
	from os import system
	from os import environ
	from os import getcwd
	from tabulate import tabulate
	import pyautogui as pag
	from PIL import Image
	from PIL import ImageDraw
	import os
	from PIL import Image, ImageDraw, ImageFont
	import textwrap
	import pandas as pd
	
except Exception as e:
	print("[#] ERROR : importing modules failed ")
	print(e)
	exit()

system("")
environ["WDM_LOG_LEVEL"] = "0"
class style():
    BLACK = '\033[30m'
    RED = '\033[31m'
    GREEN = '\033[32m'
    YELLOW = '\033[33m'
    BLUE = '\033[34m'
    MAGENTA = '\033[35m'
    CYAN = '\033[36m'
    WHITE = '\033[37m'
    UNDERLINE = '\033[4m'
    RESET = '\033[0m'
    
print(style.GREEN + "[#] Modules imported ...")

current_directory = getcwd()
# extracting the data from the xlsx sheet in data folder
try:
	data = pd.read_excel (current_directory + r'/data/data.xlsx') 
	name_list = data["NAME"].tolist()
	absent = data["absent"].tolist()
	leave = data["leave"].tolist()
	present = data["present"].tolist()
	percentage = data["percentage"].tolist()
	numbers = data["number"].tolist()

except Exception as e:
	print(style.RED + "[#] ERROR : error extracting data from data.xlsx")
	print("[#] check the file : data.xlsx in the data folder")
	print(e)
	exit()

print(style.GREEN + "[#] data loaded ...")



# creating the attendance cards
# loading the font 
print(style.WHITE + "[#] loading font file from font folder " )
try : 
	font = ImageFont.truetype(current_directory + "/fonts/Roboto-Medium.ttf",17)
	font_name = ImageFont.truetype(current_directory + "/fonts/Roboto-Black.ttf",21)
except Exception as e:
	print(style.RED + "[#] ERROR : error while loading the font file ")
	print("[#] check the font files : fonts/Roboto-Medium.ttf")
	print("[#]                        fonts/Roboto-Black.ttf")
	print(e)
	exit()
print(style.GREEN + "[#] font files loaded")
# generating the cards
try:
	black_color = (0, 0, 54)
	for name in name_list:
			img = Image.open(current_directory + '/data/bg.png')
			d = ImageDraw.Draw(img)
			idx = name_list.index(name)
			abs = absent[idx]
			lv = leave[idx]
			psnt = present[idx]
			prsnt = percentage[idx]
			
			
			d.text((60,25), str(name),fill=black_color, font=font_name)
			d.text((100,120), ("OP absent"), font=font)
			d.text((300,120), str(abs), font=font)
			d.text((100,140), ("OP leave"), font=font)
			d.text((300,140), str(lv), font=font)
			d.text((100,180), ("OP percentage"), font=font)
			d.text((300,180), str(prsnt)+" %", font=font)
	
			save_location = current_directory +"/attendance_cards/" +str(name)+".png"
			img.save(save_location, "PNG")
			print(style.WHITE + "[#] card created for " + str(name))
except Exception as e :
	print(style.RED + "[#] ERROR : error creating the attendance cards")
	print(e)
print(style.GREEN + "[#] cards created...")

# sending the certificated to the curresponding persons
print(style.CYAN + "[#] opening Google chrome to send the cards to the recipients" + style.RESET)
try:
	options = Options()
	options.add_experimental_option("excludeSwitches", ["enable-logging"])
	options.add_argument("--profile-directory=Default")
	options.add_argument("--user-data-dir=/var/tmp/chrome_user_data")
	driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
except Exception as e:
	print(style.RED + "[#] ERROR: error opening Google Chrome")
	print("[#] Google Chrome may not be installed in the system")
	print(e)
print(style.CYAN + '[#] Once your browser opens up sign in to web whatsapp')
print("[#] opening whatsapp web")
try:
	driver.get('https://web.whatsapp.com')
except Exception as e:
	print(style.RED + "[#] ERROR : error opening WhatsappWeb")
	print("[#] make sure your system is connected to the internet")
	print(e)
input(style.MAGENTA + "AFTER logging into Whatsapp Web is complete and your chats are visible, press ENTER..." + style.RESET)

