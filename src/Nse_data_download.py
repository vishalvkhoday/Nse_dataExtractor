'''
Created on Dec 20, 2018

@author: vkhoday
'''
import unittest
import ObjectRepo as Obj_r
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from HTMLTestRunner import HTMLTestRunner
import HtmlTestRunner
from datetime import date
import time

driver = Obj_r.driver


class NseDataDownload(unittest.TestCase):

    def setUp(self):
        driver.get(Obj_r.nse_Del_url)  
        driver.implicitly_wait(5)
        driver.maximize_window()
        print ("Web page loaded & Window maximized ")  
        driver.find_element_by_xpath(Obj_r.x_rd_bt_duration).click()
        pic_dt_from =driver.find_element_by_xpath(Obj_r.x_cal_From_Dt)
        # pic_dt_from.send_keys('08-09-2018')
        pic_dt_from.send_keys('01-01-2019')
        end_Dt= str(date.today()).split('-')
        pic_dt_to = driver.find_element_by_xpath(Obj_r.x_cal_To_Dt)
        pic_dt_to.send_keys(end_Dt[2]+'-'+end_Dt[1]+'-'+end_Dt[0])
        driver.find_element_by_xpath(Obj_r.x_txt_symbol).click()    
        pass


    def testNSE(self):
        row_cnt =1
        lst_script = Obj_r.get_Script_name()
        Ws = Obj_r.get_Excel_Obj()
        scr_row = Ws.max_row
        for script,scr_status in lst_script.items():
            row_cnt+=1
            if scr_status =='Yes':
                try:
#                     driver.find_element_by_xpath(Obj_r.x_txt_symbol).send_keys(script)
#                     driver.find_element_by_xpath(Obj_r.x_bt_GetData).click()
#                     driver.set_page_load_timeout(4)
#                     driver.implicitly_wait(2)
#                     time.sleep(3)
#                     driver.find_element_by_xpath(Obj_r.x_lk_dt_file).click() 
#                     Ws['C' +str(row_cnt) ]= str('No')
                    Obj_r.update_Value(row_cnt)
                    print ('{}) {} download completed remaining {}'.format(row_cnt-1,script,scr_row -row_cnt-1))            
#                     Obj_r.save_Excel()               
                    driver.find_element_by_xpath(Obj_r.x_txt_symbol).clear()
                except:
                    print ('{}) {} download failed'.format(row_cnt-1,script))
                    continue

        pass

    def tearDown(self):
        unittest.TestCase.tearDown(self)
        driver.close()
        
        pass



if __name__ == "__main__":
    #import sys;sys.argv = ['', 'Test.testName']
    unittest.main(testRunner = HtmlTestRunner.HTMLTestRunner(output='../Reports'))