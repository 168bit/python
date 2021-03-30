import json
from logging import exception
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
import time
import os
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys

#posting
def posting(email,pwd,post_id,grops):
    options = Options()
    point = "Loading Chrome..."
    try:
        options.binary_location=os.getcwd()+"\\config\\chrome-win\\chrome.exe"
        #options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-notifications')
        point = "Loading Driver..."
        driver = webdriver.Chrome(executable_path=os.getcwd()+"\\config\\chromedriver", options=options)
        point = "Loading Facebook..."
        driver.get("https://facebook.com")
        driver.find_element_by_id("email").send_keys(email)
        driver.find_element_by_id("pass").send_keys(pwd)
        driver.find_element_by_name("login").click()
        time.sleep(1)
        driver.get(post_id)
        time.sleep(3)
        i = 1
        for i in range(1,100):
            point = "Share"
            driver.find_element_by_xpath("//div[@aria-label='Send this to friends or post it on your timeline.']").click()
            time.sleep(3)
            point = "Share Group"
            driver.find_element_by_xpath("//div[. = 'Share to a group']").click()
            time.sleep(3)
            if i == 1:
                try: 
                    point = "Select Profile"
                    driver.find_element_by_xpath("//*[@class='hv4rvrfc o8rfisnq']").click()
                    time.sleep(1)
                    driver.find_element_by_xpath("//div/div[1]/div/div[4]/div/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div/div/div[1]/div/div[1]/div/div[1]").click()
                    time.sleep(1)
                except:
                    print("Not have page")
            point = "Select Group"
            driver.find_element_by_xpath("//*[@class='rq0escxv mkhogb32 pedkr2u6 jb3vyjys ph5uu5jm qt6c0cv9 b3onmgus hzruof5a pmk7jnqg kwrap0ej kr520xx4 enuw37q7 z1801hqc art1omkt nw2je8n7 hhz5lgdu']").send_keys(Keys.END)
            driver.find_element_by_xpath("//div/div["+str(i)+"]/div/div[1]/div[2]/div[1]/div/div/div[2]/span[.='Public group']").click()
            time.sleep(3)
            point = "Post"
            driver.find_element_by_xpath("//span[.='Post']").click()
            print("Post: "+str(i))
            time.sleep(3)       
            
    except exception as e:
        print(point)
        raise e
        
        time.sleep(20)
 
def load_cre(file):
    print("Loading Config...")
    try:
        with open(file) as f:
            data = f.read().splitlines()
        user = data[0][5::]
        pwd = data[1][5::]
        return (user,pwd)
    except Exception as e:
        print("Can't fix Creden.txt")

url = input("Input Post URL: ")
path = os.getcwd()+"\\config\\creden.txt"
cre = load_cre(path)
usr = cre[0]
pwd = cre[1]
posting(usr,pwd,url,1)
time.sleep(1000)

