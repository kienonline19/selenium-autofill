import pandas as pd
import time

from selenium import webdriver
from selenium.webdriver import ActionChains, Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait


def handle_dropdownlist_div(wd, input_xpath, close_btn_xpath, data):
    elem = wd.find_element(By.XPATH,
                           input_xpath)

    actions = ActionChains(wd)
    actions.move_to_element(elem).perform()

    close_btn = WebDriverWait(wd, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, close_btn_xpath))
    )

    close_btn.click()

    elem.send_keys(Keys.CONTROL + "a")
    elem.send_keys(Keys.DELETE)

    elem.send_keys("Tỉnh " + data)

    elem.send_keys(Keys.ARROW_DOWN)
    elem.send_keys(Keys.ENTER)


def xu_ly_hanh_kiem_hoc_luc(wd, div_xpath, data):
    lop_hl = wd.find_element(By.XPATH, div_xpath)
    lop_hl.click()

    hl_option = wd.find_element(By.XPATH, f"//li[text()='{data}']")
    hl_option.click()


def handle_row(wd: webdriver.Chrome, df_row):
    cccd = wd.find_element(By.XPATH,
                           "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[1]/div[2]/div/div[1]/div/div/input")
    cccd.send_keys(df_row.CCCD)

    ho_va_tendem = df_row.HỌ + ' ' + df_row.tên
    ho_ten = wd.find_element(By.XPATH,
                             "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[1]/div[2]/div/div[2]/div/input")
    ho_ten.send_keys(ho_va_tendem.upper())

    ten = browser.find_element(By.XPATH,
                               "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[1]/div[2]/div/div[3]/div/input")
    ten.send_keys(df_row.tên.upper())

    gioi_tinh = browser.find_element(By.XPATH,
                                     "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[1]/div[2]/div/div[4]/div/div")
    gioi_tinh.click()

    ul_element = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[3]/ul"))
    )

    li_element = WebDriverWait(ul_element, 10).until(
        EC.element_to_be_clickable((By.XPATH, f".//li[text()='{df_row._5.title()}']"))
    )
    li_element.click()

    ngay_sinh = wd.find_element(By.XPATH,
                                "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[1]/div[2]/div/div[5]/div/input")

    day, month, year = df_row._7.split('/')
    day = day.zfill(2)
    month = month.zfill(2)
    year = year.zfill(4)
    date = f"{day}/{month}/{year}"

    ngay_sinh.send_keys(date)

    ten_tinh = df_row.Tỉnh.title()

    handle_dropdownlist_div(wd,
                            "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[1]/div[2]/div/div[6]/div/div/input",
                            "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[1]/div[2]/div/div[6]/div/div/div/button[1]",
                            ten_tinh
                            )

    dan_toc = wd.find_element(By.XPATH,
                              "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[1]/div[2]/div/div[9]/div/div/input")
    dan_toc.send_keys(df_row._12.title())

    dan_toc.send_keys(Keys.ARROW_DOWN)
    dan_toc.send_keys(Keys.ENTER)

    handle_dropdownlist_div(
        wd,
        "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[2]/div[1]/div/div/input",
        "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[2]/div[1]/div/div/div/button[1]",
        ten_tinh
    )

    quan_huyen = wd.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[2]/div[2]/div/div/input")
    quan_huyen.send_keys("Huyện " + df_row._10.title())

    quan_huyen.send_keys(Keys.ARROW_DOWN)
    quan_huyen.send_keys(Keys.ENTER)

    time.sleep(2)

    xa = WebDriverWait(wd, 10).until(EC.presence_of_element_located(
        (By.XPATH, "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[2]/div[3]/div/div/input"))
    )
    xa.send_keys("Xã " + df_row._9.title())

    xa.send_keys(Keys.ARROW_DOWN)
    xa.send_keys(Keys.ENTER)

    # thong tin nguoi giam ho
    giam_ho = wd.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[4]/div[1]/div/input")
    giam_ho.send_keys(df_row._13.capitalize())

    truong_thcs = wd.find_element(By.XPATH,
                                  "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[7]/div[1]/div/div/input")
    truong_thcs.send_keys(df_row._15.title())

    truong_thcs.send_keys(Keys.ARROW_DOWN)
    truong_thcs.send_keys(Keys.ENTER)

    # lop
    time.sleep(2)
    lop = wd.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[7]/div[2]/div/div")
    lop.click()

    li_tag = wd.find_element(By.XPATH, "/html/body/div[2]/div[3]/ul/li[2]")
    li_tag.click()

    # nam
    from datetime import datetime as dt
    year = dt.now().year
    nam_tn = wd.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[7]/div[3]/div/input")
    nam_tn.send_keys(str(year))

    # xep loai
    xl = wd.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[7]/div[4]/div/div")
    xl.click()

    xl_text = df_row._16.capitalize()
    xl_option = wd.find_element(By.XPATH, f"//li[text()='{xl_text}']")
    xl_option.click()

    # xu ly hanh kiem - hoc luc
    xu_ly_hanh_kiem_hoc_luc(wd,
                            "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[8]/div[1]/div/div[1]/div/div",
                            df_row.HL6.capitalize()
                            )

    xu_ly_hanh_kiem_hoc_luc(wd,
                            "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[8]/div[1]/div/div[2]/div/div",
                            df_row.HK6.capitalize()
                            )

    xu_ly_hanh_kiem_hoc_luc(wd,
                            "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[8]/div[2]/div/div[1]/div/div",
                            df_row.HL7.capitalize()
                            )

    xu_ly_hanh_kiem_hoc_luc(wd,
                            "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[8]/div[2]/div/div[2]/div/div",
                            df_row.HK7.capitalize()
                            )

    xu_ly_hanh_kiem_hoc_luc(wd,
                            "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[8]/div[3]/div/div[1]/div/div",
                            df_row.HL8.capitalize()
                            )

    xu_ly_hanh_kiem_hoc_luc(wd,
                            "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[8]/div[3]/div/div[2]/div/div",
                            df_row.HK8.capitalize()
                            )

    xu_ly_hanh_kiem_hoc_luc(wd,
                            "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[8]/div[4]/div/div[1]/div/div",
                            df_row.HL9.capitalize()
                            )

    xu_ly_hanh_kiem_hoc_luc(wd,
                            "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[8]/div[4]/div/div[2]/div/div",
                            df_row.HK9.capitalize()
                            )

    time.sleep(2)
    tbm = wd.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[9]/div[1]/div/input")
    tbm.send_keys(str(df_row.TBM).replace('.', ','))

    toan_cc = wd.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[9]/div[2]/div/input")
    toan_cc.send_keys(str(df_row.TOÁN).replace('.', ','))

    van_cc = wd.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[9]/div[3]/div/input")
    van_cc.send_keys(str(df_row.VĂN).replace('.', ','))

    time.sleep(2)

    nv1 = wd.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/div[5]/div/div[2]/div[10]/div[3]/div/div/div/input")
    nv1.send_keys("trường thpt trần hưng đạo".upper())

    nv1.send_keys(Keys.ARROW_DOWN)
    nv1.send_keys(Keys.ENTER)

    nop_hs = wd.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/div[5]/div/div[1]/div/button")
    nop_hs.click()


def load_page():
    browser.get(url)

    dkts = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/div[2]/div[1]/div/div[3]/div/button[4]"))
    )

    dkts.click()


df = pd.read_excel("TTT10nhap.xlsx")

df["NGÀY SINH"] = df["NGÀY SINH"].astype(str) + '/' + df["Unnamed: 7"].astype(str) + '/' + df["Unnamed: 8"].astype(str)
df = df.drop(["Unnamed: 7", "Unnamed: 8"], axis=1)

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
browser = webdriver.Chrome(options=chrome_options)

url = "https://daknong.tuyensinhdaucap.com/admission/25a9adf7-ab1c-4caf-8368-654ab32c9e40"
load_page()

for i, row in enumerate(df.itertuples(), start=1):
    try:
        handle_row(browser, row)
        time.sleep(5)
        if i != len(df):
            load_page()
    except Exception as e:
        print(e)

time.sleep(5)
browser.quit()
