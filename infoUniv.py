# 개별대학 항목 추출 로직
import time
import sys
import openpyxl as pyxl
import keyboard
import pyperclip
import os.path

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime
from info_func import excel_insert_check, browser_handle_quit, iaif_path_append

# 파라미터
v_year = int(input("공시년도 : "))
# v_dtYear = 2018  # 기준년도 : 공시에 따라 info_year 또는 info_year-1
print("생성된 윈도우 창에서 테이블 목록 중 하나를 선택하세요.")
uList = __import__("infoUnivClickList")
time.sleep(0.5)
v_table = uList.check_us_name
# v_table = input("공시 테이블명 : ")  # 파이썬 파일명을 그대로 입력
v_table = v_table.lower()
if v_table[4:5] == "5":
    univ_grp_type = "A"
elif v_table[4:5] == "7":
    univ_grp_type = "C"
else:
    univ_grp_type = "B"
v_univ_info_except = False

v_download_folder = r"V:\document\정보공시데이터\PYTHON\excel\down_file2"
v_file_path = v_download_folder + "\\"  # 엑셀 다운로드 경로

time.sleep(0.33)
chrome_option = webdriver.ChromeOptions()
prefs = {"download.default_directory": v_download_folder}
chrome_option.add_experimental_option("prefs", prefs)
browser = webdriver.Chrome(options=chrome_option)
browser.maximize_window()
time.sleep(1)

browser_handle_quit(browser)

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
            time.sleep(0.7)
            sel_element[0].click()
            v_check_btn = True
        else:
            v_check_btn = False
        time.sleep(0.3)


# 대학 공시 검색 --start
def univ_search(u_code, u_nm, r_no):
    browser.execute_script("document.body.style.zoom='90%'")
    browser.get("http://academyinfo.go.kr/popup/pubinfo1690/list.do?schlId="+str(u_code))
    time.sleep(0.5)
    page_flow(browser.find_elements_by_xpath("//li[@class='ui-tab']/a[text()='공시정보']"))
    time.sleep(0.33)
    page_flow(browser.find_elements_by_xpath("//li[@class='ui-tab']/a[contains(text(), '전체목록')]"))
    time.sleep(0.33)
    # 공시항목 클릭
    iaif_check = False
    v_insert_check = True
    excel_text = ""
    v_count2 = 0
    if v_iaif_path:
        level_iaif = []
        for idx, val in enumerate(v_iaif_path):
            if idx == 0:
                level_iaif = browser.find_elements_by_xpath("//span[@class='text-list-title'][contains(text(), '" +
                                                            val + "')]")
            elif idx == len(v_iaif_path) - 1:
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
                try:
                    element = element[0].find_element_by_xpath("./following-sibling::button[contains(text(), '" +
                                                               str(v_year) + "')]")
                    v_insert_check = True
                    element.click()
                except NoSuchElementException:
                    v_insert_check = False
                    excel_text = "현재년도 미공시"
                break
            if v_count2 == 12:
                sys.exit()

            v_count2 += 1
            time.sleep(0.33)

    time.sleep(0.33)

    # 해당년도 항목을 찾았으면 다음 로직을 수행하고, 찾지 못했으면 다음 대학으로 이동
    if v_insert_check:
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
            time.sleep(0.33)

        datacheck = browser.find_elements_by_xpath(
            "//div[@class='textitem UbiHTMLViewer_previewpage_1font_1 UbiHTMLViewer_previewpage_1color_f_0']")
        if len(datacheck) == 0:
            datacheck = browser.find_elements_by_xpath(
                "//div[@class='textitem UbiHTMLViewer_previewpage_1font_0_0 UbiHTMLViewer_previewpage_1color_f_0_0']")

        if len(datacheck) == 1 and "없습니다" in datacheck[0].text:
            v_insert_check = False

        print("데이터 수 : "+str(len(datacheck)))

        # 모든 페이지를 불러왔는지를 확인
        v_check = False
        v_count3 = 0
        while not v_check:
            page_check = browser.find_elements_by_xpath("//div[@id='UbiHTMLViewerUbiToolbarButton_TotalPageText']")[
                0].text
            if "+" not in page_check:
                v_check = True
            elif v_count3 >= 540:
                print("모든 페이지를 불러오지 못했습니다. 확인 바랍니다.")
                wb.close()
                browser.quit()
                sys.exit()

            v_count3 += 1
            time.sleep(0.33)

        global v_save_file_name
        v_save_file_name = ""
        if len(datacheck) > 1:
            time.sleep(0.5)
            page_flow(browser.find_elements_by_xpath("//td[@id='UbiHTMLViewerUbiToolbar_SaveButton']/input"))  # 저장버튼 클릭
            time.sleep(2)
            v_count3 = 0
            v_check = False
            this_year = datetime.today().year
            while not v_check:  # 60초
                v_save_file_list = os.listdir(v_file_path)
                if len(v_save_file_list) > 0:
                    for idx, file in enumerate(v_save_file_list):
                        if str(this_year) in file and not file.endswith(".crdownload"):
                            v_save_file_name = os.listdir(v_file_path)[idx]
                            v_check = True
                            print(v_save_file_name + " saved")
                            time.sleep(0.6)
                            os.rename(v_file_path + v_save_file_name, v_file_path + u_nm + ".xlsx")
                            break

                v_count3 += 1
                if v_count3 > 40:
                    print(v_iaif_name + "이 다운로드 되지 않았거나 여러 파일이 존재합니다.")
                    wb.close()
                    browser.quit()
                    sys.exit()
                time.sleep(0.3)

        browser.close()
        browser.switch_to.window(browser.window_handles[0])

    excel_insert_check(ws, r_no, v_insert_check, excel_text)
    time.sleep(0.33)
# 대학 공시 검색 --end


v_file_add_text = "_전문대학" if univ_grp_type == "C" else ""
file_path = "V:\document\정보공시데이터\PYTHON\excel\\"
wb = pyxl.load_workbook(file_path+v_iaif_name+v_file_add_text+".xlsx")
# wb = pyxl.load_workbook(file_path+"개별대학 추출 대학 목록.xlsx")
global ws
ws = wb['Sheet1']  # or .active

for r in ws.rows:
    idx = r[0].row
    if idx > 1 and r[3].value is None:
        univ_nm = r[0].value
        print(univ_nm)
        univ_id = r[1].value
        univ_search(univ_id, univ_nm, idx)
        wb.save(file_path+v_iaif_name+v_file_add_text+".xlsx")
    # if idx == 3:
    #     break
wb.close()
browser.quit()
sys.exit()
