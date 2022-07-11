# 공시항목의 칼럼 목록을 불러오는 파일
import time
import sys

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

from info_func import get_col_title, element_check, iaif_path_append


# 파라미터
info_year = input("공시년도 : ")
# v_dtYear = 2018  # 기준년도 : 공시에 따라 info_year 또는 info_year-1
info_table = input("테이블명(소문자) : ")  # 소문자로 입력 : 파이썬 파일명을 그대로 찾아야 함
info_table = info_table.lower()
univ_grp_type = ""  # 태그의 option value - 01[전문대학], 02[대학], 03[대학원]
if info_table[4:5] == "5":
    univ_grp_type = "02"
elif info_table[4:5] == "7":
    univ_grp_type = "01"
else:
    univ_grp_type = "03"
v_univ_info_col_add = True

# part_list = ["A", "B", "C", "D", "E"]
part_name = {"A": "인문ㆍ사회계열", "B": "자연과학계열", "C": "예ㆍ체능계열", "D": "공학계열", "E": "의학계열"}
v_part_string = 'B'

browser = webdriver.Chrome()
browser.maximize_window()
time.sleep(0.66)

browser.switch_to.window(browser.window_handles[1])
browser.close()
browser.switch_to.window(browser.window_handles[0])

v_col_univ_info = [
 "학교종류"
,"설립구분"
,"지역"
,"상태"
]

# 공시 파일 동적 import
iaif_path_append()
v_iaif_name = __import__(info_table).iaif_name

v_iaif_path = None
try:
    v_iaif_path = __import__(info_table).iaif_path
except AttributeError:
    pass

v_iaif_ord_name = None
try:
    v_iaif_ord_name = __import__(info_table).iaif_ord_name
except AttributeError:
    pass

# 공시항목 크롤링 처리 로직
browser.get("http://academyinfo.go.kr/index.do")
time.sleep(1.5)
if len(browser.window_handles) == 2:
    browser.switch_to.window(browser.window_handles[1])
    browser.close()
    browser.switch_to.window(browser.window_handles[0])

soup = BeautifulSoup(browser.page_source, 'html.parser')


# 페이지 내 클릭에 의한 다음 루트까지의 요소 검색 지연 처리
def page_flow(browser_element):
    v_count2 = 0
    v_check_btn = False
    while not v_check_btn:
        sel_element = browser_element
        if sel_element:
            sel_element[0].click() if isinstance(sel_element, list) else sel_element.click()
            v_check_btn = True
        else:
            v_check_btn = False

        if v_count2 == 13:
            print("요소 찾기 시간 초과")
            sys.exit()

        v_count2 += 1
        time.sleep(0.33)


if v_iaif_ord_name:
    # element[0].send_keys(v_iaif_org_name)
    browser.get('http://academyinfo.go.kr/search/search.do?kwd="' + v_iaif_ord_name + '"&schlKnd=' + univ_grp_type)
else:
    # element[0].send_keys(v_iaif_name)
    browser.get('http://academyinfo.go.kr/search/search.do?kwd="' + v_iaif_name + '"&schlKnd=' + univ_grp_type)

# element[0].send_keys(Keys.ENTER)
time.sleep(1.5)

# 공시 데이터 검색
global element
try:
    element = element_check(browser.find_element_by_xpath(
        "//table[@class='tbl-col']/tbody[@id='targetDiv']//td[contains(text(),'" +
        v_iaif_name + "')]"))
except:
    try:
        element = element_check(browser.find_element_by_xpath(
            "//table[@class='tbl-col']/tbody[@id='targetDiv']//td[contains(text(),'" +
            v_iaif_ord_name + "')]"))
    except:
        print("공시 항목을 찾을 수 없습니다.")
        sys.exit()

# 공시명 우측 '학교별평균값' select box 클릭
actions = ActionChains(browser)
# xpath상에서 index는 1부터 시작
actions.move_to_element(element.find_element_by_xpath("..//span[contains(@class, 'ui-select-wrap')][1]"))
actions.click()
actions.perform()
time.sleep(0.33)

# 학과별 클릭
page_flow(browser.find_element_by_xpath("//div[contains(@class, 'ui-selectmenu-open')]//div[text()='학과별']"))

# 계열 select box 선택하기
page_flow(element.find_element_by_xpath("..//span[@id='pgmStags']"))

# 계열 선택하기
page_flow(browser.find_element_by_xpath(
    "//div[contains(@class, 'ui-selectmenu-open')]//div[text()='" + part_name[v_part_string] + "']"))

# 년도 클릭
page_flow(element.find_element_by_xpath("..//button[@data-svy_yr='" + str(info_year) + "']"))

browser.switch_to.window(browser.window_handles[1])
bs_soup = BeautifulSoup(browser.page_source, 'html.parser')

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

print(get_col_title(browser))
browser.quit()
sys.exit()
