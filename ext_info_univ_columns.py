# 개별대학 공시항목의 칼럼 목록을 불러오는 파일
import time
import sys
import openpyxl as pyxl
import keyboard
import pyperclip

from bs4 import BeautifulSoup
from selenium import webdriver

from info_func import get_col_title, iaif_path_append

# 파라미터
v_year = int(input("공시년도 : "))
# v_dtYear = 2018  # 기준년도 : 공시에 따라 info_year 또는 info_year-1
v_table = input("공시 파일명(소문자) : ")  # 소문자로 입력 : 파이썬 파일명을 그대로 입력
v_table = v_table.lower()
v_univnm = input("대학명[ALL: ENTER] : ")  # 특정 대학을 지정하면 해당 대학 다운로드

v_univgrp1 = ""  # 태그의 option value - 01[전문대학], 02[대학], 03[대학원]
if v_table[4:5] == "5":
    univ_grp_type = "3"
elif v_table[4:5] == "7":
    univ_grp_type = "4"
else:
    univ_grp_type = "5"
v_univ_info_except = False

time.sleep(0.33)
chrome_option = webdriver.ChromeOptions()
prefs = {"download.prompt_for_download": True}
chrome_option.add_experimental_option("prefs", prefs)
browser = webdriver.Chrome(options=chrome_option)
browser.maximize_window()
time.sleep(0.66)

browser.switch_to.window(browser.window_handles[1])
browser.close()
browser.switch_to.window(browser.window_handles[0])

v_col_univ_info = [
    "학교종류"
    , "설립구분"
    , "지역"
    , "상태"
]

# 공시 파일 내 변수 가져오기
iaif_path_append()
v_mod = __import__(v_table)

v_total_col = v_mod.total_col
v_iaif_name = v_mod.iaif_name
print(v_iaif_name)
v_iaif_path = None
try:
    v_iaif_path = __import__(v_table).iaif_path
except AttributeError:
    pass

# 대학원 취업 현황의 전문대학원과 특수대학원은 따로 공시되지만 같은 공시 테이블을 사용하므로
# 뒤에 오는 특수대학원에 공시 테이블명을 두어 다른 파일명이지만 같은 테이블에 insert 되도록 처리
try:
    v_table = v_mod.table_name
except AttributeError:
    pass

try:
    v_dtyear_idx = v_mod.dtyear_idx
except AttributeError:
    v_dtyear = v_year - v_mod.dtyear_num
    v_dtyear_idx = -1  # 공시를 조회한 항목에서 넣어야 할 때, 그 외는 dtyear_num값 만큼 평가년도에서 감소


# 공시항목 크롤링 처리 로직
# browser.get("http://academyinfo.go.kr/index.do")

# soup = BeautifulSoup(browser.page_source, 'html.parser')


# 페이지 내 클릭에 의한 다음 루트까지의 요소 검색 지연 처리
def page_flow(browser_element):
    v_check_btn = False
    while not v_check_btn:
        sel_element = browser_element
        if sel_element:
            sel_element[0].click()
            v_check_btn = True
        else:
            v_check_btn = False
        time.sleep(0.33)


# 대학 공시 검색 --start
def univ_search(u_id):
    browser.get("http://academyinfo.go.kr/popup/pubinfo1690/list.do?schlId="+str(u_id))

    page_flow(browser.find_elements_by_xpath("//li[@class='ui-tab']/a[text()='공시정보']"))
    page_flow(browser.find_elements_by_xpath("//li[@class='ui-tab']/a[contains(text(), '전체목록')]"))
    time.sleep(0.5)
    # 공시항목 클릭
    iaif_check = False
    v_count2 = 0
    if v_iaif_path:
        level_iaif = []
        for f_idx, val in enumerate(v_iaif_path):
            if f_idx == 0:
                level_iaif = browser.find_elements_by_xpath("//span[@class='text-list-title'][contains(text(), '" +
                                                            val + "')]")
            elif f_idx == len(v_iaif_path) - 1:
                level_iaif = level_iaif[0].find_element_by_xpath("../ul//span[contains(text(), '" + val + "')]")
            else:
                level_iaif = level_iaif[0].find_element_by_xpath("../ul//span[contains(text(),'" + val + "')]")
            time.sleep(0.33)
        level_iaif = level_iaif.find_element_by_xpath("./following-sibling::button[contains(text(), '" +
                                                      str(v_year) + "')]")
        level_iaif.click()
    else:
        while not iaif_check:
            # depth-text 클래스 span 태그로 먼저 찾음. 없으면 상위 태그인 text-list-title 클래스 span 태그를 찾는다.
            element = browser.find_elements_by_xpath("//span[@class='depth-text'][contains(text(), '" +
                                                     v_iaif_name + "')]")
            if not element:
                element = browser.find_elements_by_xpath("//span[@class='text-list-title'][contains(text(), '" +
                                                         v_iaif_name + "')]")
            if element:
                element = element[0].find_element_by_xpath("./following-sibling::button[contains(text(), '" +
                                                           str(v_year) + "')]")
                element.click()
                break
            if v_count2 == 12:
                sys.exit()

            v_count2 += 1
            time.sleep(0.33)

    time.sleep(0.33)
    browser.switch_to.window(browser.window_handles[1])

    # 공시 항목 데이터 창 open 후, 데이터 불러오기 까지 지연 시키기
    v_check = False
    v_count3 = 0
    while not v_check and v_count3 <= 540:  # 540 : 약 3분
        bs_soup = BeautifulSoup(browser.page_source, 'html.parser')
        check_elmt = bs_soup.find("div", {"id": "UbiHTMLViewer_preview_1"})
        if check_elmt:
            v_check = True

        v_count3 += 1
        time.sleep(0.35)

    print(get_col_title(browser))

    browser.quit()
    sys.exit()
# 대학 공시 검색 --end


file_path = r"V:\document\정보공시데이터\PYTHON\excel\\"
wb = pyxl.load_workbook(file_path+"대학 목록.xlsx")
ws = wb['Sheet1']  # or .active


def excel_insert_check(row_id):
    ws.cell(row=row_id, column=4).value = "O"


for r in ws.rows:
    idx = r[0].row
    v_stop_chk = False
    if idx > 1:
        if r[0].value == v_univnm or v_univnm == "":
            univ_id = r[1].value
            univ_search(univ_id)
            v_stop_chk = True
        else:
            continue
    if v_stop_chk:
        break
browser.quit()
wb.close()
sys.exit()
