import sys
from ExcelEditor import ExcelEditor
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as cond
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

data = []
total_points = 0
current_points = 0

class DataObj():
    def __init__(self, word, weight):
        self.word = word
        self.weight = weight

class AnalyzeStatement():
    def __init__(self):
        pass

    def create_driver(self):
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--window-size=1920x1080")
        return webdriver.Chrome(chrome_options=chrome_options)

    def close_driver(self, driver):
        driver.quit()
    
    def get_data(self, file, sheet_name):
        total = 0
        excel = ExcelEditor(file)
        row_count = excel.get_row_count(sheet_name)
        for i in range(row_count):
            row_values = excel.get_row(sheet_name, i)
            word = row_values[0]
            weight = row_values[1]
            total += weight
            global data
            data.append( DataObj(word, weight) )
        global total_points
        total_points = total
        for obj in data:
            print(obj.word + ' ' + str(obj.weight))
        
    def search_google(driver, statement):
        driver.get("https://www.google.com/")
        driver.find_element_by_css_selector()

    # def search_twitter(driver, statement):

    def calculate_points():
        outcome = ""
        calculation = current_points - total_points
        if calculation < 0 and calculation > -0.02:
            outcome = "neutral"







if __name__ == '__main__':
    analyze = AnalyzeStatement()
    if len(sys.argv) != 2:
        print("Please enter one statement.")
    else:
        analyze.get_data('data.xlsx','Sheet1')
        driver = analyze.create_driver()
        search_google(driver, sys.argv[1])



