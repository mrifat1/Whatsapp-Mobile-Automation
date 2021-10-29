import time
from appium import webdriver
import pandas as pd
FILE_LOC = "data.xlsx"
SRNO = 'SRNO'
NAME = 'Name'
PHONENUMBER = 'Phone Number'
MESSAGE = 'Message'
data = []
def read_data_from_excel():

  df = pd.read_excel(FILE_LOC)

  for i in df.index:
    number = str(df[PHONENUMBER][i])
    output = {
      'SrNo': df[SRNO][i],
      'Name': df[NAME][i],
      'PhoneNumber': number,
      'Message': df[MESSAGE][i],
      'Message_check': df['Message_check'][i]
    }
    data.append(output)

def whatsappmob():
  desired_cap={
    "deviceName": "Android Emulator",
    "platformName": "Android",
    "noReset":'true',
    "appPackage": "com.whatsapp",
    "appActivity": "com.whatsapp.HomeActivity",

  }
  df = pd.read_excel(FILE_LOC)
  driver = webdriver.Remote("http://localhost:4723/wd/hub",desired_cap)
  for files in data:
    search_button = driver.find_element_by_id("com.whatsapp:id/menuitem_search")
    search_button.click()

    search_box = driver.find_element_by_id("com.whatsapp:id/search_input")
    search_box.send_keys(str(files['PhoneNumber']))

    Boro_pera_for_entering_to_the_text_box = driver.find_element_by_xpath("/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.FrameLayout[2]/android.widget.LinearLayout/androidx.recyclerview.widget.RecyclerView/android.widget.RelativeLayout/android.widget.LinearLayout")
    Boro_pera_for_entering_to_the_text_box.click()
    time.sleep(2)
    text_box = driver.find_element_by_id("com.whatsapp:id/entry")
    text_box.send_keys(str(files['Message']))
    time.sleep(2)
    send_button = driver.find_element_by_id("com.whatsapp:id/send")
    send_button.click()
    time.sleep(2)
    back_button = driver.find_element_by_id("com.whatsapp:id/whatsapp_toolbar_home")
    back_button.click()
    df['Message_check'] = ('Sent')
    df.to_excel('new.xlsx')

if __name__ == '__main__':

    read_data_from_excel()
    whatsappmob()

