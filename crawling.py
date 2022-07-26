# -*- coding: utf-8-sig -*-
from PyQt5 import QtCore, QtGui, QtWidgets, uic
import pandas as pd
import sys
import numpy as np
from PIL import Image
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import os
import time
import math
from selenium import webdriver
from PyQt5.QtWidgets import QApplication
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains 
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
# import chromedriver_autoinstaller
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.support.select import Select


options = webdriver.ChromeOptions()
options.add_argument("start-maximized");
options.add_argument("disable-infobars")
options.add_argument("--disable-extensions")
options.add_argument("--incognito")

def resource_path(relative_path): 
    """ Get absolute path to resource, works for dev and for PyInstaller """ 
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__))) 
    return os.path.join(base_path, relative_path)

form = resource_path('ui.ui') 

form_class = uic.loadUiType(form)[0]


class app_class(QMainWindow, form_class) :
    
    def __init__(self) :
        super().__init__()
        self.setFixedSize(760, 400)
        self.setupUi(self)
        self.save_path_find.clicked.connect(self.save_dirct_open)
        self.craw_start.clicked.connect(self.search)
        self.setWindowTitle("g2b_crawling")
    
        QApplication.processEvents()
    
    def save_dirct_open(self) :
        global save_fd_nm
        save_fd_nm = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Directroy") + "/"
        self.save_path.setPlainText(save_fd_nm)

    def search(self) : 
        global key_word
        key_word = self.key_word.toPlainText()
        
        if  getattr(sys, 'frozen', False): 
            # chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
            driver = webdriver.Chrome(ChromeDriverManager().install())
        else:
            driver = webdriver.Chrome(ChromeDriverManager().install())
        
        if self.real_anno.isChecked() == True :
            while(True) : 
                try:
                    # 입찰정보 검색 페이지로 이동
                    driver.get('https://www.g2b.go.kr:8101/ep/tbid/tbidFwd.do')
                    
                    # 업무 종류 체크
                       
                    #task_dict = {'용역': 'taskClCds5'}
                    #for task in task_dict.values():
                        #checkbox = driver.find_element_by_id(task)
                        #checkbox.click()
                    
                    # 검색어
                    query = key_word
                    
                    # 공고명에 해당하는 태그 가져오기
                    bidNm = driver.find_element_by_id('bidNm')
                    
                    # 내용을 삭제 (혹시 뭔가 입력되어 있을 수도 있으니 날리고 시작해야 함)
                    bidNm.clear()
                    
                    # 검색어 입력후 엔터
                    bidNm.send_keys(query)
                    bidNm.send_keys(Keys.RETURN)
                    
                    # 검색 조건 체크
                    option_dict = {'검색기간 1달': 'setMonth1_1', '입찰마감건 제외': 'exceptEnd', '검색건수 표시': 'useTotalCount'}
                    for option in option_dict.values():
                        checkbox = driver.find_element_by_id(option)
                        checkbox.click()
                                   
                    # 공고종류 선택 (드롭다운에서 전체 선택)
                    driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[3]/form/table/tbody/tr[12]/td[1]/div/select/option[1]').click()
                    
                    # 목록수 선택 (드롭다운에서 100건 선택)
                    recordcountperpage = driver.find_element_by_name('recordCountPerPage')
                    selector = Select(recordcountperpage)
                    selector.select_by_value('100')
                
                    # 검색 버튼 클릭
                    search_button = driver.find_element_by_class_name('btn_mdl')
                    search_button.click()
                    
                    # 검색 결과 확인
                    elem = driver.find_element_by_class_name('results')
                    div_list = elem.find_elements_by_tag_name('div')
                    
                    # 검색 결과 모두 긁어서 리스트로 저장
                    results = []
                    for div in div_list:
                        results.append(div.text)
                        
                    # 검색결과 모음 리스트를 12개씩 분할하여 새로운 리스트로 저장 
                    result = [results[i * 10:(i + 1) * 10] for i in range((len(results) + 10 - 1) // 10 )]
                    
                    # 결과 엑셀파일로 저장
                    col_list = ['업무','공고번호/차수','분류','공고명','공고기관','수요기관','계약방법','입력일시','공동수급','투찰']
                    result_final = pd.DataFrame(result, columns = col_list)
                    df = pd.DataFrame(columns = col_list)
                    df = pd.concat([df, result_final], ignore_index= True)
                    df.to_excel(save_fd_nm + "데이터관련사업_본공고"+ ".xlsx", index = False, encoding = 'utf-8-sig')
                    model = DataFrameModel(df)
                    self.csv_sample.setModel(model)
                    break
                    
                    
                except Exception as e:
                    # 위 코드에서 에러가 발생한 경우 출력
                    print(e)
        elif self.pre_anno.isChecked() == True :
            while(True) : 
                try:
                    # 입찰정보 검색 페이지로 이동
                    driver.get('https://www.g2b.go.kr:8081/ep/preparation/prestd/preStdSrch.do?taskClCd=5')
                    # 검색어
                    query = key_word
                    
                    #  품명(사업명) 해당 태그 가져오기
                    prodNm = driver.find_element_by_id('prodNm')    
                
                    # 내용을 삭제 (혹시 뭔가 입력되어 있을 수도 있으니 날리고 시작해야 함)
                    prodNm.clear()
                    
                    # 검색어 입력후 엔터
                    prodNm.send_keys(query)
                    prodNm.send_keys(Keys.RETURN)                
                
                    # 목록수 선택 (드롭다운에서 100건 선택)
                    recordcountperpage = driver.find_element_by_name('recordCountPerPage')
                    selector = Select(recordcountperpage)
                    selector.select_by_value('100')
                
                    # 검색 버튼 클릭
                    search_button = driver.find_element_by_class_name('btn_mdl')
                    search_button.click()
                    
                    # 검색 결과 확인
                    elem = driver.find_element_by_class_name('results')
                    div_list = elem.find_elements_by_tag_name('div')
                    
                    # 검색 결과 모두 긁어서 리스트로 저장
                    results = []
                    for div in div_list:
                        results.append(div.text)
                        
                    # 검색결과 모음 리스트를 7개씩 분할하여 새로운 리스트로 저장 
                    result = [results[i * 7:(i + 1) * 7] for i in range((len(results) + 7 - 1) // 7 )]
                    
                    # 결과 엑셀파일로 저장
                    col_list = ['No','등록번호','참조번호','품명(사업명)','수요기관','사전규격공개일시','업체등록의견수']
                    result_final = pd.DataFrame(result, columns = col_list)
                    df = pd.DataFrame(columns = col_list)
                    df = pd.concat([df, result_final], ignore_index= True)
                    df.참조번호 = df.참조번호.astype(str)
                    df.to_excel(save_fd_nm + "데이터관련사업_사전규격"+ ".xlsx", index = False, encoding = 'utf-8-sig')
                    model = DataFrameModel(df)
                    self.csv_sample.setModel(model)        
                    break
                
                except Exception as e:
                    # 위 코드에서 에러가 발생한 경우 출력
                    print(e)

    

class DataFrameModel(QtCore.QAbstractTableModel):
    DtypeRole = QtCore.Qt.UserRole + 1000
    ValueRole = QtCore.Qt.UserRole + 1001

    def __init__(self, df=pd.DataFrame(), parent=None):
        super(DataFrameModel, self).__init__(parent)
        self._dataframe = df

    def setDataFrame(self, dataframe):
        self.beginResetModel()
        self._dataframe = dataframe.copy()
        self.endResetModel()

    def dataFrame(self):
        return self._dataframe

    dataFrame = QtCore.pyqtProperty(pd.DataFrame, fget=dataFrame, fset=setDataFrame)

    @QtCore.pyqtSlot(int, QtCore.Qt.Orientation, result=str)
    def headerData(self, section: int, orientation: QtCore.Qt.Orientation, role: int = QtCore.Qt.DisplayRole):
        if role == QtCore.Qt.DisplayRole:
            if orientation == QtCore.Qt.Horizontal:
                return self._dataframe.columns[section]
            else:
                return str(self._dataframe.index[section])
        return QtCore.QVariant()

    def rowCount(self, parent=QtCore.QModelIndex()):
        if parent.isValid():
            return 0
        return len(self._dataframe.index)

    def columnCount(self, parent=QtCore.QModelIndex()):
        if parent.isValid():
            return 0
        return self._dataframe.columns.size

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < self.rowCount() \
            and 0 <= index.column() < self.columnCount()):
            return QtCore.QVariant()
        row = self._dataframe.index[index.row()]
        col = self._dataframe.columns[index.column()]
        dt = self._dataframe[col].dtype

        val = self._dataframe.iloc[row][col]
        if role == QtCore.Qt.DisplayRole:
            return str(val)
        elif role == DataFrameModel.ValueRole:
            return val
        if role == DataFrameModel.DtypeRole:
            return dt
        return QtCore.QVariant()

    def roleNames(self):
        roles = {
            QtCore.Qt.DisplayRole: b'display',
            DataFrameModel.DtypeRole: b'dtype',
            DataFrameModel.ValueRole: b'value'
        }
        return roles      
        

        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = app_class()
    myWindow.show()
    app.exec_()
