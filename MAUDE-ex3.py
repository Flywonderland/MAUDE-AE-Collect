from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
import xlwt
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options

class AEsearch(object):
    def __init__(self):
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        prefs = {"profile.managed_default_content_settings.images": 2}
        chrome_options.add_experimental_option("prefs", prefs)
        self.dr =webdriver.Chrome(chrome_options=chrome_options)
        self.dr.get('https://www.accessdata.fda.gov/scripts/cdrh/cfdocs/cfMAUDE/search.CFM')

    
    def search(self):
        Manu = input("Manu")
        Brand = input("Brand")
        Fromdate = input("From")
        Todate = input("To")
        form = self.dr.find_element_by_id("mdr-form")
        manu = form.find_element_by_id("Manufacturer").send_keys(Manu)
        brand = form.find_element_by_xpath('form/table/tbody/tr/td/input[@name="BrandName"]').send_keys(Brand)
        form.find_element_by_name("ReportDateFrom").clear()
        fromdate = form.find_element_by_name("ReportDateFrom").send_keys(Fromdate)
        form.find_element_by_name("ReportDateTo").clear()                                                                 
        todate = form.find_element_by_name("ReportDateTo").send_keys(Todate)
        bnt = form.find_element_by_name("Search").click()

        MaudeData=[]
        
        element = ['Catalog number','Lot number','Model number','Is This An Adverse Event Report','Is This A Product Problem Report','Webaddress','ReportNumber','MDR Report Key']
        MaudeData.append(element)
        
        try:
            flag = self.dr.find_element_by_name("submaudeform")
        except NoSuchElementException as e:
            flag = False
        else:
            flag = True

        if flag == True:

            content = self.dr.find_element_by_id("user_provided")
            result = content.find_element_by_xpath('table/tbody/tr/td/b')
            print("返回结果："+result.text)
            p =int(result.text)
            a = (p-1)//10

            
            for i in range(0,a+1):
                local = self.dr.current_window_handle
                self.dr.switch_to_window(local)
                Lists = self.dr.find_elements_by_xpath('//*[@id="user_provided"]/table[3]/tbody/tr/td/table/tbody/*/td[2]/a')
            
                for List in Lists:
                    
                    URL = List.get_attribute("href")
                    
                    js = "window.open('about:blank')"
                    self.dr.execute_script(js)
                    all_handles = self.dr.window_handles

                    for handle in all_handles:

                        if handle != local:
                            self.dr.switch_to_window(handle)
                            self.dr.get(URL)
                            time.sleep(1)

                            data = []

                            try:
                                CatNo = self.dr.find_element_by_xpath('//*[@id="user_provided"]/table/tbody/tr/td/table/tbody/tr/td/*/*/*/th[contains(text(),"Device Catalogue Number")]/following-sibling::td')
                            except NoSuchElementException as e:
                                data.append('Not Found')
                            else:
                                data.append(CatNo.text)

                            try:
                                LoNo = self.dr.find_element_by_xpath('//*[@id="user_provided"]/table/tbody/tr/td/table/tbody/tr/td/*/*/*/th[contains(text(),"Device LOT Number")]/following-sibling::td')
                            except NoSuchElementException as e:
                                data.append('Not Found')
                            else:
                                data.append(LoNo.text)

                            try:
                                MoNo = self.dr.find_element_by_xpath('//*[@id="user_provided"]/table/tbody/tr/td/table/tbody/tr/td/*/*/*/th[contains(text(),"Device MODEL Number")]/following-sibling::td')
                            except NoSuchElementException as e:
                                data.append('Not Found')
                            else:
                                data.append(MoNo.text)

                            try:
                                AE = self.dr.find_element_by_xpath('//*[@id="user_provided"]/*/*/tr/td/table/tbody/tr/td/*/*/*/th[contains(text(),"Is This An Adverse Event Report")]/following-sibling::td')
                            except NoSuchElementException as e:
                                data.append('Not Found')
                            else:
                                data.append(AE.text) 
             
                            try:
                                DP = self.dr.find_element_by_xpath('//*[@id="user_provided"]/*/*/tr/td/table/tbody/tr/td/*/*/*/th[contains(text(),"Is This A Product Problem Report")]/following-sibling::td')
                            except NoSuchElementException as e:
                                data.append('Not Found')
                            else:
                                data.append(DP.text)

                            data.append(URL)

                            try:
                                ReportNo = self.dr.find_element_by_xpath('//*[@id="user_provided"]/*/*/tr/td/table/tbody/tr/td/*/*/*/th[contains(text(),"Report Number")]/following-sibling::td')
                            except NoSuchElementException as e:
                                data.append('Not Found')
                            else:
                                data.append(ReportNo.text)

                            try:
                                MDR = self.dr.find_element_by_xpath('//*[@id="user_provided"]/*/*/tr/td/table/tbody/tr/td/*/*/*/th[contains(text(),"MDR Report Key")]/following-sibling::td')
                            except NoSuchElementException as e:
                                data.append('Not Found')
                            else:
                                data.append(MDR.text)
                                
                                
                            self.dr.close()
                            self.dr.switch_to_window(local)

                            MaudeData.append(data)

                if i == a:
                    continue
                else:
                    try:
                        nextpage = self.dr.find_element_by_xpath('//*[@id="user_provided"]/table/*/*/*/*/*/*/*/*/*/*/*/*/table/tbody/tr/td/a[@title="Next"]')
                        nextpage.click()
                    except NoSuchElementException as e:
                        nextpage = self.dr.find_element_by_xpath('//*[@id="user_provided"]/table[2]/tbody/tr[1]/td[2]/div/table/tbody/tr/td[8]/table/tbody/tr[1]/td/table/tbody/tr/td/a')
                        nextpage.click()
        else:
            print("No result found")

        workbook =xlwt.Workbook(encoding = 'ascii')
        
        localtime=time.strftime("%Y-%m-%d %H：%M：%S", time.localtime()) 
        worksheet = workbook.add_sheet(Brand+' '+localtime)

        for r in range(0,p+1):
            for l in range(0,8):
                worksheet.write(r,l, label= MaudeData[r][l])

        workbook.save('Excel_AEsearch.xls')
        self.quit()
            
    def quit(self):
        self.dr.quit()


AEsearch().search()


