# 개별대학 항목 추출 로직
import time
import sys
import re
import cx_Oracle
import os.path
import ctypes

from bs4 import BeautifulSoup
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from info_func import get_col_title, get_insert_text, browser_handle_quit, iaif_path_append

os.environ["NLS_LANG"] = ".AL32UTF8"  # DB 케릭터셋과 일치시킴


wb, ws, v_cell_row, browser = "", "", 0, ""
dsn, conn = "", ""
result, insert_data, column_length = "", "", ""
v_year, v_table, univ_grp_type, v_mod = "", "", "", ""
v_cols_order = ""  # 항목 칼럼이 상이할 시 col, col2, ...에서 지정하는 순서
v_manual_insert_chk = False

# 파라미터
v_year = int(input("공시년도 : "))
v_table = input("공시 파일명 : ")  # 소문자로 입력 : 파이썬 파일명을 그대로 입력
v_table = v_table.lower()

iaif_path_append()
v_mod = __import__(v_table)

v_iaif_name = v_mod.iaif_name
v_iaif_ref_name = v_mod.iaif_ref_name

try:
    v_dtyear_idx = v_mod.dtyear_idx
except AttributeError:
    v_dtyear = v_year - v_mod.dtyear_num
    v_dtyear_idx = -1  # 공시를 조회한 항목에서 넣어야 할 때, 그 외는 dtyear_num값 만큼 평가년도에서 감소

v_iaif_path = None
try:
    v_iaif_path = __import__(v_table).iaif_path
except AttributeError:
    pass

browser = webdriver.Chrome()
browser.maximize_window()
time.sleep(1)

browser_handle_quit(browser)

# oracle 연동
dsn = cx_Oracle.makedsn("61.81.234.137", 1521, "COGDW")
conn = cx_Oracle.connect("dusd", "dusd$#@!", dsn)
cursor = conn.cursor()

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


def current_item_chk(browser_text, v_check_name):
    v_count2 = 0
    v_check_btn = False
    while not v_check_btn:
        sel_element = browser_text
        if sel_element == v_check_name:
            v_check_btn = True
        else:
            v_check_btn = False

        if v_count2 == 10:
            print("요소 비교 시간 초과")
            sys.exit()

        v_count2 += 1
        time.sleep(0.33)


def insert_logic():  # 테이블 insert 처리
    # print(insert_data)
    # show_data(insert_data)
    v_table_cols = ""
    # print('col ord : '+v_cols_order)
    if v_cols_order == "":
        v_table_cols = """(""" + ",".join(v_mod.table_cols) + """)""" if v_mod.table_cols else ""
    elif v_cols_order == "2":
        v_table_cols = """(""" + ",".join(v_mod.table_cols2) + """)""" if v_mod.table_cols2 else ""
    elif v_cols_order == "3":
        v_table_cols = """(""" + ",".join(v_mod.table_cols3) + """)""" if v_mod.table_cols3 else ""

    insert_sql = """INSERT INTO """ + v_table + v_table_cols + """ VALUES (""" + get_insert_text(column_length) + \
                 v_mod.insert_last_col + """)"""
    cursor_ins = conn.cursor()
    try:
        # insert_data = unicode(insert_data, "euc-kr").encode("utf-8")
        cursor_ins.executemany(insert_sql, insert_data)
        cursor_ins.close()

        cursor.execute("UPDATE IAIF_US_DL_CHECK SET DL_YN = 'Y', DT_CREA = SYSDATE" +
                       " WHERE IAIF_NM = UPPER('" + v_iaif_ref_name + "')" +
                       "   AND INFO_YYYY = '" + str(v_year) + "'" +
                       "   AND DHW_UNIV_ID = " + str(dhw_u_id))
        conn.commit()
        print(u_name + " ------ insert success!")
    except cx_Oracle.DatabaseError as e:
        error, = e.args
        print(insert_sql)
        print("error code : " + str(error.code))
        print(error.message)
        print(error.context)
        cursor_ins.close()
        conn.rollback()


# browser.switch_to.window(browser.window_handles[1])
# browser.close()
# browser.switch_to.window(browser.window_handles[0])

cursor.execute("SELECT UNIV_CODE, REPLACE(DHW_UNIV_NM,'_본교','') AS UNIV_NM, DHW_UNIV_NM, DHW_UNIV_ID" +
               "  FROM IAIF_US_DL_CHECK A" +
               " WHERE IAIF_NM = UPPER('" + v_iaif_ref_name + "')" +
               "   AND INFO_YYYY = '" + str(v_year) + "'" +
               "   AND (DL_YN <> 'Y' OR DL_YN IS NULL)"
               " ORDER BY A.UNIV_NM, A.UNIV_ID, A.DHW_UNIV_NM")
result = cursor.fetchall()
bf_u_id, v_element = 0, ""
v_trans_char, v_change_char = "·・·∙", "ㆍㆍㆍㆍ"
v_trans_char2 = v_trans_char + " "

for idx, row in enumerate(result):
    print(row)
    u_id = row[0]
    u_name = row[1]
    dhw_u_name = row[2]
    dhw_u_id = row[3]

    if bf_u_id != u_id:
        bf_u_id = u_id
        # 해당 본교 대학원 접속
        browser.get("http://academyinfo.go.kr/popup/pubinfo1690/list.do?schlId=" + bf_u_id)

    # 해당 대학원 이름 찾기
    try:
        v_element = browser.find_element_by_xpath("//ul[@class='college-info-list']//a[translate(text(), '" +
                                                  v_trans_char2 + "', '" + v_change_char + "')='" +
                                                  u_name.replace(" ", "") + "']")
        v_element.location_once_scrolled_into_view
        page_flow(v_element)
    except NoSuchElementException:
        try:
            u_name = u_name + "(폐교)"
            v_element = browser.find_element_by_xpath("//ul[@class='college-info-list']//a[translate(text(), '" +
                                                      v_trans_char2 + "', '" + v_change_char + "')='" +
                                                      u_name.replace(" ", "") + "']")
            v_element.location_once_scrolled_into_view
            page_flow(v_element)
        except NoSuchElementException:
            print("해당 대학원을 찾을 수 없습니다.")
            cursor.close()
            conn.close()
            sys.exit()

    current_item_chk(browser.find_element_by_xpath("//ul[@class='college-info-list']/li[@class='current']/a").text
                     .translate(str.maketrans(v_trans_char, v_change_char, " ")), u_name.replace(" ", ""))

    v_element = browser.find_element_by_xpath("//li[contains(@class, 'ui-tab')]/a[text()='공시정보']")
    v_element.location_once_scrolled_into_view
    page_flow(v_element)
    current_item_chk(browser.find_element_by_xpath("//li[contains(@class, 'ui-tabs-active')]/a").text, "공시정보")
    page_flow(browser.find_elements_by_xpath("//div[@class='ui-tabs sub-tabs']//li[@class='ui-tab']/a[contains(text(), '전체목록')]"))
    current_item_chk(browser.find_element_by_xpath(
        "//div[@class='ui-tabs sub-tabs']//li[@class='ui-tab ui-tabs-active']/span[contains(text(), '전체목록')]/..")
                     .get_attribute("class"), "ui-tab ui-tabs-active")

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

    # 공시 데이터 처리
    browser.switch_to.window(browser.window_handles[1])
    global bs_soup
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

    # 현재 공시와 저장된 공시의 칼럼 비교 : ext_info_columns.get_col_title
    cur_title_list = get_col_title(browser)

    try:
        if cur_title_list == v_mod.total_col:
            v_cols_order = ""
        elif cur_title_list == v_mod.total_col2:
            v_cols_order = "2"
        elif cur_title_list == v_mod.total_col3:
            v_cols_order = "3"
    except AttributeError:
        ctypes.windll.user32.MessageBoxW(0, str(v_year) + "년도 공시의 제목 열이 기존 제목 열과 다릅니다. 확인 바랍니다.", "", 4096)
        conn.close()
        sys.exit()
    # ###

    # 모든 페이지를 불러왔는지를 확인
    v_check = False
    v_count3 = 0
    while not v_check:
        page_check = browser.find_elements_by_xpath("//div[@id='UbiHTMLViewerUbiToolbarButton_TotalPageText']")[
            0].text
        if "+" not in page_check:
            v_check = True
        elif v_count3 >= 540:
            ctypes.windll.user32.MessageBoxW(0, str(v_year) + "모든 페이지를 불러오지 못했습니다. 확인 바랍니다.", "", 4096)
            conn.close()
            sys.exit()

        v_count3 += 1
        time.sleep(0.33)

    scroll_element = "UbiHTMLViewer_previewframe"
    scrollElem = browser.find_element(By.ID, value=scroll_element)

    docHeight = browser.execute_script("return document.scrollingElement.scrollHeight")
    scrollHeight = browser.execute_script("return document.getElementById('" + scroll_element + "').scrollHeight")

    actions = ActionChains(browser)
    actions.move_to_element(scrollElem)
    actions.click()
    actions.perform()

    curHeight = 0
    i = 1
    while curHeight < scrollHeight:
        curHeight = docHeight * i
        browser.execute_script("document.getElementById('" + scroll_element + "').scrollTop = " + str(curHeight))
        time.sleep(1.2)
        i += 1

    bs_soup2 = BeautifulSoup(browser.page_source, 'html.parser')
    myElem3 = bs_soup2.find("div", {"id": scroll_element})

    insert_data = []  # table insert data

    column_length = 0  # insert 행의 컬럼 길이

    my_element4 = bs_soup2.find("div", {"id": "UbiHTMLViewer_previewpage_1"})
    my_element4 = my_element4.find_all("div", {"class": "UbiHTMLViewer_previewpage_1color_b_2"})

    v_univ_title_index = None
    v_minus_val = 0
    for idx, val in enumerate(my_element4):
        if idx == 0 and ("연도" in val.get_text() or "년도" in val.get_text()):
            v_minus_val = 1
        if val.get_text() in ["학교명", "학교", "대학명", '대학원명']:
            v_univ_title_index = idx - v_minus_val
            break


    # 정보공시 대학명 검사
    def is_univ_nm_check(val):
        univ_nm_str = ["대학", "학교", "기술원", "본교", "캠퍼스", "분교"]
        gbn = False
        for str in univ_nm_str:
            if str in val:
                gbn = True
        return gbn


    # 정보공시 대학원명 대학교 뒤 공백 검사
    def is_univ_nm_space_check(val):
        univ_nm_str = ["대학교", "기술원"]
        univ_txt = ""
        for str in univ_nm_str:
            if str in val:
                univ_txt = str

        string_idx = val.index(univ_txt)
        if val[string_idx + len(univ_txt)] != " ":
            return val.replace(univ_txt, univ_txt + " ")
        else:
            return val


    for child in myElem3.children:
        div_elmt = child.findAll("div", {"class": "textitem"})
        univ_data = {}
        for idx, el in enumerate(div_elmt):
            attr_style = el.attrs["style"]
            if el.parent.get("id") + "color_b_2" not in el.attrs["class"] and el.parent.get(
                    "id") + "color_b_2_0" not in el.attrs["class"]:
                # 페이지 당 칼럼 제목과 기준년도 데이터는 제외(기타)
                re_left = re.compile("left: [0-9]{1,4}px;+")  # style에서 left 추출(px제외)
                re_top = re.compile("top: [0-9]{1,4}px;+")  # style에서 top 추출(px제외)
                re_width = re.compile("width: [0-9]{1,4}px;+")  # style에서 width 추출(px제외)
                css_left = re_left.search(attr_style)
                css_top = re_top.search(attr_style)
                css_width = re_width.search(attr_style)
                css_left = css_left.group().replace("px;", "").split(" ")[1]
                css_top = css_top.group().replace("px;", "").split(" ")[1]
                css_width = css_width.group().replace("px;", "").split(" ")[1]

                if int(css_width) >= 1000:
                    continue

                # left와 top 스타일 수치를 이용하여 dictionary 생성
                def css_grouping(v_css_left, v_css_top):
                    data = ""
                    # 대학명이면 "_" 이전 공백 제거 또는 "_"가 없는 대학교명의 맨 끝 공백 제거, 수치이면 천단위 콤마 제거
                    if is_univ_nm_check(el.get_text()):
                        data = el.get_text().strip().replace(" _", "_")
                    elif v_table in ("iaif5200_13", "iaif6200", "iaif7200_13", "iaif5341", "iaif5541"):
                        # 졸업생 취업률의 취업률 칼럼은 16년 이전에는 '-'로 표시되었기 때문에 ''로 수정
                        data = "" if el.get_text() == "-" else el.get_text().replace(",", "")
                    else:
                        # 수치 데이터의 ','(comma) 표시 제거
                        data = el.get_text().replace(",", "")  # .replace("∙", "ㆍ")

                    # if v_univgrp1 == "03":
                    #     data = is_univ_nm_space_check(el.get_text())

                    if css_top in univ_data:
                        univ_data[v_css_top][v_css_left] = data
                    else:
                        univ_data[v_css_top] = {v_css_left: data}


                if int(css_top) > 55:
                    if v_dtyear_idx > -1 or int(css_left) >= 0:  # 국가 칼럼 포함
                        css_grouping(css_left, css_top)

        row_list = []
        bf_list = []
        univ_data = sorted(univ_data.items(), key=lambda x: int(x[0]))
        for key, val in univ_data:
            # top(key)별로 값({left: data, ...})을 left순으로 정렬
            sort_item = sorted(val.items(), key=lambda x: int(x[0]))
            item_list = list(list(zip(*sort_item))[1])
            key_list = list(list(zip(*sort_item))[0])

            # style left 값으로 기준 list를 만들고, 이것으로 각 행을 비교하여 병합된 부분을 이전 행에서 메꾼다.
            if not row_list:
                row_list = key_list[:]
            elif len(row_list) != len(key_list):
                for idx, val in enumerate(row_list):
                    if idx == len(key_list) or val != key_list[idx]:
                        key_list.insert(idx, val)
                        item_list.insert(idx, bf_list[idx])

            bf_list = item_list[:]  # 이전 리스트를 다음 행에서 참조하기 위해 할당함

            # 합계 첫 행 패스
            if item_list[0].replace(" ", "") == "합계":
                continue

            # 특정 공시의 경우 기준년도가 첫 열이 아닌 다른 열에 존재할 수도 있기 때문에 이를 아래와 같이 처리함
            insert_year = [str(v_year)]
            try:
                if v_mod.dtyear_idx:
                    pass
            except AttributeError:
                insert_year.append(str(v_dtyear))

            # None 데이터를 지정된 데이터로 대체하기 위함. item_list는 기준년도 칼럼 제외되어 있음. 주의 v_mod.data_modify[0]
            try:
                if v_mod.data_modify:
                    v_data_modify_list = v_mod.data_modify
                    for idx, val in enumerate(item_list):
                        if idx == v_data_modify_list[0] and val == "":
                            item_list[idx] = v_data_modify_list[1]
            except AttributeError:
                pass

            append_data = insert_year + [dhw_u_name] + item_list + [dhw_u_id]

            insert_data.append(append_data)

            if len(insert_data) == 1:
                # print(insert_data[0])
                column_length = len(insert_data[0])

    insert_logic()
    browser.close()
    browser.switch_to.window(browser.window_handles[0])


cursor.close()
conn.close()
browser.quit()
sys.exit()
