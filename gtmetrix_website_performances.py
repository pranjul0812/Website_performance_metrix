import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *


class WebsitePerformanceCalculator:

    def read_excel(self):
        df = pd.read_excel("C:\\Users\\Pranjul Mishra\\Desktop\\URL_list.xlsx")
        df.drop(df.columns[0], axis=1, inplace=True)
        header = df.iloc[1]
        new_df = pd.DataFrame(df.values[2:], columns=header)
        print(new_df)
        return new_df

    def write_excel(self, df, path=""):
        if path:
            filename = path + "/GtMetrix_website_performances.xlsx"
        else:
            filename = "GtMetrix_website_performances.xlsx"
        df.to_excel(filename, index=False, header=True)

    def gtmetrix(self, driver, df, wait):
        gt_output_time = []
        gt_output_grade = []
        driver.maximize_window()
        driver.get("https://gtmetrix.com/")
        driver.implicitly_wait(3)

        for index, row in df.iterrows():
            data = row['Website URL']
            if pd.notnull(data):
                try:
                    tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='url']")))
                    submit = driver.find_element(By.XPATH, "//div[@class='analyze-form-button']//button[@type='submit']")
                    tab.clear()
                    time.sleep(2)
                    tab.send_keys(data)
                    submit.click()
                    elements = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[@class='report-score']")))
                    grade_element = elements[0].find_element(By.XPATH, ".//i[contains(@class, 'grade')]")
                    grade = grade_element.get_attribute('class')
                    gt_output_grade.append(grade[-1])
                    time_element = driver.find_element(By.XPATH,
                                                       "//div[@class='report-page-detail'][1]//span[@class='report-page-detail-value']")
                    time_value = time_element.text
                    gt_output_time.append(time_value)
                    print(grade[-1], " ", time_value)
                    driver.back()
                    time.sleep(2)
                except TimeoutException:
                    gt_output_time.append("ERROR")
                    gt_output_grade.append("ERROR")
                    print("Error", " ", "Error")
                    driver.back()
            else:
                gt_output_time.append("0")
                gt_output_grade.append("0")
                print("No Grade", "No Time value")
        df['gt_metrix_grade'] = gt_output_grade
        df['gt_metrix_time'] = gt_output_time

        self.write_excel(df)


sa = WebsitePerformanceCalculator()
driverLocation = "C:\\Pranjul\\webdriver\\chromedriver.exe"
os.environ['webdriver.driver.chrome'] = driverLocation
driver = webdriver.Chrome(driverLocation)
wait = WebDriverWait(driver, 120, poll_frequency=2)
df = sa.read_excel()
sa.gtmetrix(driver, df, wait)