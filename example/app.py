import browser
import time

Name = input("Please enter your Google Search?")

ie = browser.ie_browser()
ie.navigate('https://www.google.com/')

ie.wait_page()
ie.ie.Document.documentElement.getElementsByClassName("gLFyf gsfi")[0].value = Name
ie.click_button(name='btnK')
ie.wait_page()
time.sleep(1)
for i in ie.send_command('getElementsByClassName("g")'):
    print(i.getElementsByTagName("h3")[0].innerText)
