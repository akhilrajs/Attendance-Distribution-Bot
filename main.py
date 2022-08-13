print("[#] Attendance Distribution BOT ")
print("[#] author : https://github.com/akhilrajs")
print("")
print("")
print("[#] importing modules ...")


try:
    import requests
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options 
    from selenium.webdriver.support.ui import WebDriverWait 
    from selenium.webdriver.support import expected_conditions as EC 
    from selenium.webdriver.common.by import By 
    from webdriver_manager.chrome import ChromeDriverManager 
    from time import sleep 
    from os import system 
    from os import environ 
    from os import getcwd 
    from os import startfile 
    import pyautogui as pag 
    from PIL import Image 
    from PIL import ImageDraw 
    from PIL import ImageFont 
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

print(style.GREEN + "[#] downloading xpaths" + style.RESET)
sleep(0.8)
try :
	url = "https://raw.githubusercontent.com/akhilrajs/Attendance-Distribution-Bot/main/xpath/click_btn.txt"
	data = requests.get(url)
	click_btn_xpath = data.text
except Exception as e:
	print(style.RED + "[#] ERROR --> " + str(e) )
	print(style.RED + "[#] connect to the internet and try again " + style.RESET) 
	exit()

try :
	url = "https://raw.githubusercontent.com/akhilrajs/Attendance-Distribution-Bot/main/xpath/menu.txt"
	data = requests.get(url)
	menu_xpath = data.text
except Exception as e:
	print(style.RED + "[#] ERROR --> " + str(e) )
	print(style.RED + "[#] connect to the internet and try again " + style.RESET) 
	exit()

try :
	url = "https://raw.githubusercontent.com/akhilrajs/Attendance-Distribution-Bot/main/xpath/crop.txt"
	data = requests.get(url)
	crop_xpath = data.text
except Exception as e:
	print(style.RED + "[#] ERROR --> " + str(e) )
	print(style.RED + "[#] connect to the internet and try again " + style.RESET) 
	exit()

try :
	url = "https://raw.githubusercontent.com/akhilrajs/Attendance-Distribution-Bot/main/xpath/message_box.txt"
	data = requests.get(url)
	msg_box_xpath = data.text
except Exception as e:
	print(style.RED + "[#] ERROR --> " + str(e) )
	print(style.RED + "[#] connect to the internet and try again " + style.RESET) 
	exit()

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
			d.text((100,120), ("Absent"), font=font)
			d.text((300,120), str(abs), font=font)
			d.text((100,140), ("Leave"), font=font)
			d.text((300,140), str(lv), font=font)
			d.text((100,180), ("Percentage"), font=font)
			d.text((300,180), str(prsnt)+" %", font=font)
	
			save_location = current_directory +"/attendance_cards/" +str(name)+".png"
			img.save(save_location, "PNG")
			print(style.WHITE + "[#] card created for " + str(name))
except Exception as e :
	print(style.RED + "[#] ERROR : error creating the attendance cards")
	print(e)
	exit()
print(style.GREEN + "[#] cards created...")

# sending the certificated to the curresponding persons
print(style.CYAN + "[#] opening Google chrome to send the cards to the recipients" + style.RESET)
try:
	options = Options()
	options.add_experimental_option("excludeSwitches", ["enable-logging"])
	options.add_argument("--profile-directory=Default")
	options.add_argument("--user-data-dir=C:\\User\\Data\\Default")
	driver = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=options)
except Exception as e:
	print(style.RED + "[#] ERROR: error opening Google Chrome")
	print("[#] Google Chrome may not be installed in the system")
	print(e)
	exit()
print(style.CYAN + '[#] Once your browser opens up sign in to web whatsapp')
print("[#] opening whatsapp web" + style.RESET)
try:
	driver.get('https://web.whatsapp.com')
except Exception as e:
	print(style.RED + "[#] ERROR : error opening WhatsappWeb")
	print(style.RED + "[#] make sure your system is connected to the internet" + style.RESET)
	print(e)
	exit()
delay = 60
menu = WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.XPATH, menu_xpath)))
				

# making splitscreen
pag.hotkey('win','right')
pag.moveTo(0,1078)
file = startfile(current_directory + "\\attendance_cards")
sleep(5)
pag.hotkey('win','left')
pag.moveTo(0,1078)
pag.moveTo(322,0)
pag.click()
pag.keyDown("up")
pag.keyDown("down")
pag.keyDown("up")
pag.hotkey('ctrl','c')

fail = []
total_number = len(numbers)
j=0
for idx, number in enumerate(numbers): 
    pag.moveTo(322,0) 
    pag.click() 
    number = str(number).strip() 
    if number == "": 
        continue 
    print(style.YELLOW + '{}/{} => Sending message to {}.'.format((idx+1), total_number, number) + style.RESET) 
    try: 
        sleep(2) 
        msg = ""
        msg = "*Attendance Report* : " + str(name_list[idx]).title()  
        url = 'https://web.whatsapp.com/send?phone=+91' + number + '&text=' + msg 
        sleep(2) 
        sent = False 
        for i in range(3): 
            if not sent: 
                driver.get(url) 
                try: 
                    msg_box_xpath = WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.XPATH, menu_xpath))) 
                    sleep(1)
                    pag.moveTo(1678,978) 
                    pag.click() 
                    sleep(3) 
                    pag.hotkey('ctrl','v') 
                    crop =  WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.XPATH, crop_xpath)))
					#click_btn =  WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div/span")))
					#click_btn.click() 
                    pag.keyDown("enter") 
                except Exception as e: 
                    print(style.RED + f"\nFailed to send message to: {number}, retry ({i+1}/3)") 
                    print("Make sure your phone and computer is connected to the internet.") 
                    print("If there is an alert, please dismiss it." + style.RESET) 
                    fail.append(number) 
                else:
					#click_btn.click() 
                    sleep(1) 
                    sent=True 
                    sleep(3) 
                    print(style.GREEN + 'Message sent to: ' + number + " " + name_list[idx].title() + style.RESET) 
                    pag.moveTo(322,0) 
                    pag.click() 
                    pag.keyDown("down") 
                    pag.hotkey('ctrl','c')
            
    except Exception as e: 
        print(style.RED + 'Failed to send message to ' + number + str(e) + style.RESET) 
        fail.append(number)
        pag.moveTo(322,0) 
        pag.click() 
        pag.keyDown("down") 
        pag.hotkey('ctrl','c')
print("Failed to send to : " + '\n' )
for n in fail:
	print(n)
sleep(10)
driver.close()
file.close()
