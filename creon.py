import os
import time
from pywinauto import application

from flask import Flask
import os, sys, ctypes
import win32com.client
import pandas as pd
from datetime import datetime, timedelta
from slacker import Slacker
import time, calendar
import numpy as np
import requests

# import constants
# import util

# slack = Slacker('xoxb-1558363067780-1552204452210-3vrWUR7RH81MytFwQbRnEJXe')
# def self.dbgout(message):
#     """인자로 받은 문자열을 파이썬 셸과 슬랙으로 동시에 출력한다."""
#     print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
#     strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
#     slack.chat.post_message('#lordkonig-stock', strbuf)

# os.system('taskkill /IM coStarter* /F /T')
# os.system('taskkill /IM CpStart* /F /T')
# os.system('wmic process where "name like \'%coStarter%\'" call terminate')
# os.system('wmic process where "name like \'%CpStart%\'" call terminate')
# time.sleep(5)
# app2 = application.Application()
# app2.start('D:\CREON\STARTER\coStarter.exe /prj:cp /id:insooya /pwd:Passwo1! /pwdcert:Password12! /autostart')
# time.sleep(60)


class Creon:

    def post_message(self, token, channel, text):
        response = requests.post("https://slack.com/api/chat.postMessage",
            headers={"Authorization": "Bearer "+token},
            data={"channel": channel,"text": text}
        )
        # print(response)
    
    def dbgout(self, message):
        """인자로 받은 문자열을 파이썬 셸과 슬랙으로 동시에 출력한다."""
        myToken = "xoxb-1558363067780-1552204452210-3vrWUR7RH81MytFwQbRnEJXe"
        print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
        strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
        self.post_message(myToken,"#lordkonig-stock", strbuf)

    def printlog(self, message, *args):
        """인자로 받은 문자열을 파이썬 셸에 출력한다."""
        print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message, *args)
    
    def __init__(self):
        # 크레온 플러스 공통 OBJECT
        # 참고로 크레온 API 는 다른 증권사들 API처럼 요청에 3가지의 제한이 있음
        # 1. 주문/계좌 관련: 15초에 20회 제한
        # 2. 시세 관련: 15초에 60회 제한
        # 3. 시세관련 실시간 요청: 총 400건 제한
        self.cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
        self.cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
        self.cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
        self.cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
        self.instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
        self.cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
        # cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
        # cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
        # cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')
        # cpJpbid = win32com.client.Dispatch('Dscbo1.StockJpbid2')  #★매도잔량, 매수잔량 구하기!!
        # cpJpBid1 = win32com.client.Dispatch('Dscbo1.StockJpBid')  #★매도호가 10개, 매수호가 10개 구하기!!
        # objCpSvr7326 = win32com.client.Dispatch("cpsysdib.CpSvr7326")
        # MarkCptl8548 = win32com.client.Dispatch("CpSysDib.CpSvr8548")
        self.MarkCptlEye = win32com.client.Dispatch("CpSysDib.MarketEye")









    def kill_client(self):
        os.system('taskkill /IM coStarter* /F /T')
        os.system('taskkill /IM CpStart* /F /T')
        os.system('taskkill /IM DibServer* /F /T')
        os.system('wmic process where "name like \'%coStarter%\'" call terminate')
        os.system('wmic process where "name like \'%CpStart%\'" call terminate')
        os.system('wmic process where "name like \'%DibServer%\'" call terminate')

    def connect(self, id_, pwd, pwdcert, trycnt=300):
        if not self.connected():
            self.disconnect()
            self.kill_client()
            app = application.Application()
            # app.start('D:\CREON\STARTER\coStarter.exe /prj:cp /id:insooya /pwd:Passwo1! /pwdcert:Password12! /autostart' )
            app.start('D:\CREON\STARTER\coStarter.exe /prj:cp /id:insooya /pwd:Passwo1! /pwdcert:Password12! /autostart'.format(id=id_, pwd=pwd, pwdcert=pwdcert) )
            # app.start('D:\CREON\STARTER\coStarter.exe /prj:cp /id:{id} /pwd:{pwd} /pwdcert:{pwdcert} /autostart'.format(id=id_, pwd=pwd, pwdcert=pwdcert) )
        cnt = 0
        while not self.connected():
            if cnt > trycnt:
                return False
            time.sleep(1)
            cnt += 1
        return True

    def connected(self):
        b_connected = self.cpStatus.IsConnect
        if b_connected == 0:
            return False
        return True

    def disconnect(self):
        if self.connected():
            self.cpStatus.PlusDisconnect()
            return True
        return False








    def check_creon_system(self):
        """크레온 플러스 시스템 연결 상태를 점검한다."""
        # 관리자 권한으로 프로세스 실행 여부
        if not ctypes.windll.shell32.IsUserAnAdmin():
            self.printlog('check_creon_system() : admin user -> FAILED')
            return False
    
        # 연결 여부 체크
        if (self.cpStatus.IsConnect == 0):
            self.printlog('check_creon_system() : connect to server -> FAILED')
            return False
    
        # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
        if (self.cpTradeUtil.TradeInit(0) != 0):
            self.printlog('check_creon_system() : init trade -> FAILED')
            return False
        return True

    def get_current_price(self, code):
        """인자로 받은 종목의 현재가, 매수호가, 매도호가를 반환한다."""
        self.cpStock.SetInputValue(0, code)  # 종목코드에 대한 가격 정보
        self.cpStock.BlockRequest()
        item = {}
        item['cur_price'] = self.cpStock.GetHeaderValue(11)   # 현재가
        item['ask'] =  self.cpStock.GetHeaderValue(16)        # 매수호가
        item['bid'] =  self.cpStock.GetHeaderValue(17)        # 매도호가    
        return item['cur_price'], item['ask'], item['bid']

    def get_ohlc(self, code, qty):  #★★★ OHLC: open(시가), high(고가), low(저가), close(종가)
        """인자로 받은 종목의 OHLC 가격 정보를 qty 개수만큼 반환한다."""
        self.cpOhlc.SetInputValue(0, code)            # 종목코드
        self.cpOhlc.SetInputValue(1, ord('2'))        # 1:기간(날짜), 2:개수
        #self.cpOhlc.SetInputValue(2, ToDate)         # 시작일자 - 위에서 (1, ord('1)) 로 설정할경우
        #self.cpOhlc.SetInputValue(3, FromDate)       # 끝일자 - 위에서 (1, ord('1)) 로 설정할경우
        self.cpOhlc.SetInputValue(4, qty)             # 요청개수
        self.cpOhlc.SetInputValue(5, [0,2,3,4,5, 8, 9])  #★★★0날짜,1시간(2~5:OHLC,★8거래량,9거래대금)
        self.cpOhlc.SetInputValue(6, ord('D'))        # 'D'일봉,'W'주봉,'M'월봉,★★★'m'분봉,'T'
        self.cpOhlc.SetInputValue(9, ord('1'))        # 0:무수정주가, 1:수정주가
        self.cpOhlc.BlockRequest()
        count = self.cpOhlc.GetHeaderValue(3)   # 3:수신개수(요청개수),즉 여기서는 qty수. GetHeaderValue,GetDataValue로 결과만들기
        #columns = ['open', 'high', 'low', 'close']
        columns = ['open', 'high', 'low', 'close', 'tradevol', 'trademoney']  #★★★
        index = []
        rows = []
        for i in range(count):   #★GetDataValue(type,index)는 data를 표의 형태로 가져오는것!
            index.append(self.cpOhlc.GetDataValue(0, i))  #★GetDataValue(type,index) 요청한 data의 index (type은 열을의미, 즉,여기서 type0=날짜, type1=시가)
            rows.append([self.cpOhlc.GetDataValue(1, i), self.cpOhlc.GetDataValue(2, i),
                self.cpOhlc.GetDataValue(3, i), self.cpOhlc.GetDataValue(4, i),
                self.cpOhlc.GetDataValue(5, i), self.cpOhlc.GetDataValue(6, i) ])  #★★★
        # Calculate average volume  #★★★
        #averageVolume = (sum(columns['tradevol']) - columns['tradevol'][0]) / (len(columns['tradevol']) -1)  #★★★
        #if (columns['tradevol'][0] > averageVolume * 3):  #★★★
        #    return 1  #★★★
        #else:  #★★★
        #    return 0  #★★★
        df = pd.DataFrame(rows, columns=columns, index=index)
        #df = df.sort_index(ascending=False)  # 최신 데이터가 가장 아래로 가게 정렬
        #df.to_csv("{}.csv".format(stockcode), index=False)  #★★★csv파일로 저장하기!    
        return df


    def get_todayclose(self, code):
        """금일 시가를 반환한다."""
        try:
            time_now = datetime.now()  #현재 시간 추출
            str_today = time_now.strftime('%Y%m%d')
            ohlc = self.get_ohlc(code, 10)  #코드에 해당하는 10일치 ohlc값을 가져온다
            if str_today == str(ohlc.iloc[0].name):  #10일치중 첫번째값의 name이 오늘과 똑같으면,
                today_open = ohlc.iloc[0].open  # today_open 설정
                today = ohlc.iloc[0]  # today 설정
                # lastday = ohlc.iloc[1]  #어제값은 10일치 데이터중 두번째 값 설정  #★★★ 종목이 생긴지 1일밖에 안되어서 lastday 종가값이 없을수 있어서, lastday 는 인식못하도록 삭제함!!
                if len(ohlc) <2:  #★★★★★ 종목이 생긴지 1일밖에 안되어서 lastday 가격값이 없을수 있어서, 이경우 lastday 를 today 와 동일하게 인식하도록함!!
                    lastday = ohlc.iloc[0]
                else:
                    lastday = ohlc.iloc[1]
            else:
                # lastday = ohlc.iloc[0]  #만약10일치 데이터중 첫번째 값이 오늘이 아니라면(즉 당일조회가 아니라 밤12시 지나고 장시작전 조회할 경우), 10일치중 첫번째값이 어제데이터
                # today_open = lastday[3]  #금일 시가를 전일 종가로 설정
                today = ohlc.iloc[0]  # today 설정
            # lastday_high = lastday[1]  #시고저종 설정 - 고가
            # lastday_low = lastday[2]  #시고저종 설정 - 저가
            # lastday_close = lastday[3]  #시고저종 설정 - 종가  #★★★
            today_close = today[3]  #시고저종 설정 - 종가  #★★★
            # lastday_vol = lastday[4]  #시고저종 설정 - 거래량  #★★★
            # target_price = today_open + (lastday_high - lastday_low) * 0.2   # 0.3
            if ( today[3] != 0 ) and (not today[3] >0 ):  #★★★
                return today_close  #★
            else:
                return today_close
        except Exception as ex:
#             self.dbgout("`get_todayclose() -> exception! " + str(ex) + "`")
            return None


    def get_lastclose(self, code):
        """전일 종가를 반환한다."""
        try:
            time_now = datetime.now()  #현재 시간 추출
            str_today = time_now.strftime('%Y%m%d')
            ohlc = self.get_ohlc(code, 10)  #코드에 해당하는 10일치 ohlc값을 가져온다
            if str_today == str(ohlc.iloc[0].name):  #10일치중 첫번째값의 name이 오늘과 똑같으면,
                today = ohlc.iloc[0]  # today 설정
                today_open = ohlc.iloc[0].open  # today_open 설정
                if len(ohlc) <2:  #★★★★★ 종목이 생긴지 1일밖에 안되어서 lastday 가격값이 없을수 있어서, 이경우 lastday 를 today 와 동일하게 인식하도록함!!
                    lastday = ohlc.iloc[0]
                else:
                    lastday = ohlc.iloc[1]
            else:
                today = ohlc.iloc[0]  # today 설정
                today_open = ohlc.iloc[0].open  # today_open 설정
                # today_open = lastday[3]  #금일 시가를 전일 종가로 설정
                # lastday = ohlc.iloc[1]  #만약10일치 데이터중 첫번째 값이 오늘이 아니라면(즉 당일조회가 아니라 밤12시 지나고 장시작전 조회할 경우), 10일치중 첫번째값이 어제데이터
                if len(ohlc) <2:  #★★★★★ 종목이 생긴지 1일밖에 안되어서 lastday 가격값이 없을수 있어서, 이경우 lastday 를 today 와 동일하게 인식하도록함!!
                    lastday = ohlc.iloc[0]
                else:
                    lastday = ohlc.iloc[1]
            lastday_high = lastday[1]  #시고저종 설정 - 고가
            lastday_low = lastday[2]  #시고저종 설정 - 저가
            lastday_close = lastday[3]  #시고저종 설정 - 종가  #★★★
            today_close = today[3]  #시고저종 설정 - 종가  #★★★
            #lastday_vol = lastday[4]  #시고저종 설정 - 거래량  #★★★
            #target_price = today_open + (lastday_high - lastday_low) * 0.2   # 0.3
            if ( lastday[3] != 0 ) and (not lastday[3] >0 ):  #★★★
                return lastday_close  #★
            else:
                return lastday_close
        except Exception as ex:
#             self.dbgout("`get_lastclose() -> exception! " + str(ex) + "`")
            return None


    def get_todayTrMoney(self, code):
        """금일 시가를 반환한다."""
        try:
            time_now = datetime.now()  #현재 시간 추출
            str_today = time_now.strftime('%Y%m%d')
            ohlc = self.get_ohlc(code, 10)  #코드에 해당하는 10일치 ohlc값을 가져온다
            if str_today == str(ohlc.iloc[0].name):  #10일치중 첫번째값의 name이 오늘과 똑같으면,
                today_open = ohlc.iloc[0].open  # today_open 설정
                today = ohlc.iloc[0]  # today 설정
                lastday = ohlc.iloc[1]  #어제값은 10일치 데이터중 두번째 값 설정
            else:
                lastday = ohlc.iloc[0]  #만약10일치 데이터중 첫번째 값이 오늘이 아니라면(즉 당일조회가 아니라 밤12시 지나고 장시작전 조회할 경우), 10일치중 첫번째값이 어제데이터
                today_open = lastday[3]  #금일 시가를 전일 종가로 설정
                today = ohlc.iloc[0]  # today 설정
            #lastday_high = lastday[1]  #시고저종 설정 - 고가
            #lastday_low = lastday[2]  #시고저종 설정 - 저가
            #lastday_close = lastday[3]  #시고저종 설정 - 종가  #★★★
            #today_close = today[3]  #시고저종 설정 - 종가  #★★★
            today_TrMoney = today[5]  #시고저종 설정 - 거래대금  #★★★
            #lastday_vol = lastday[4]  #시고저종 설정 - 거래량  #★★★
            #target_price = today_open + (lastday_high - lastday_low) * 0.2   # 0.3
            return today_TrMoney
        except Exception as ex:
#             self.dbgout("`get_todayTrMoney() -> exception! " + str(ex) + "`")
            return None

    def get_1daybef_TrMoney(self, code):
        """금일 시가를 반환한다."""
        try:
            time_now = datetime.now()  #현재 시간 추출
            str_today = time_now.strftime('%Y%m%d')
            ohlc = self.get_ohlc(code, 30)  #코드에 해당하는 30일치 ohlc값을 가져온다
            if len(ohlc) > 2:
                if str_today == str(ohlc.iloc[0].name):  #30일치중 첫번째값의 name이 오늘과 똑같으면,
                    today_open = ohlc.iloc[0].open  # today_open 설정
                    today = ohlc.iloc[0]  # today 설정
                    targetday = ohlc.iloc[1]  #어제값은 10일치 데이터중 두번째 값 설정
                else:
                    lastday = ohlc.iloc[0]  #만약10일치 데이터중 첫번째 값이 오늘이 아니라면(즉 당일조회가 아니라 밤12시 지나고 장시작전 조회할 경우), 10일치중 첫번째값이 어제데이터
                    today_open = lastday[3]  #금일 시가를 전일 종가로 설정
                    targetday = ohlc.iloc[1]  #어제값은 10일치 데이터중 두번째 값 설정
            else:
                targetday = ohlc.iloc[len(ohlc)-1]
            #lastday_high = lastday[1]  #시고저종 설정 - 고가
            #lastday_low = lastday[2]  #시고저종 설정 - 저가
            #lastday_close = lastday[3]  #시고저종 설정 - 종가  #★★★
            #today_close = today[3]  #시고저종 설정 - 종가  #★★★
            targetday_TrMoney = targetday[5]  #시고저종 설정 - 거래대금  #★★★
            #lastday_vol = lastday[4]  #시고저종 설정 - 거래량  #★★★
            #target_price = today_open + (lastday_high - lastday_low) * 0.2   # 0.3
            return targetday_TrMoney
        except Exception as ex:
#             self.dbgout("`get_1daybef_TrMoney() -> exception! " + str(ex) + "`")
            return None

    def get_2daybef_TrMoney(self, code):
        """금일 시가를 반환한다."""
        try:
            time_now = datetime.now()  #현재 시간 추출
            str_today = time_now.strftime('%Y%m%d')
            ohlc = self.get_ohlc(code, 30)  #코드에 해당하는 30일치 ohlc값을 가져온다
            if len(ohlc) > 3:
                if str_today == str(ohlc.iloc[0].name):  #30일치중 첫번째값의 name이 오늘과 똑같으면,
                    today_open = ohlc.iloc[0].open  # today_open 설정
                    today = ohlc.iloc[0]  # today 설정
                    targetday = ohlc.iloc[2]  #어제값은 10일치 데이터중 두번째 값 설정
                else:
                    lastday = ohlc.iloc[0]  #만약10일치 데이터중 첫번째 값이 오늘이 아니라면(즉 당일조회가 아니라 밤12시 지나고 장시작전 조회할 경우), 10일치중 첫번째값이 어제데이터
                    today_open = lastday[3]  #금일 시가를 전일 종가로 설정
                    targetday = ohlc.iloc[2]  #어제값은 10일치 데이터중 두번째 값 설정
            else:
                targetday = ohlc.iloc[len(ohlc)-1]
            #lastday_high = lastday[1]  #시고저종 설정 - 고가
            #lastday_low = lastday[2]  #시고저종 설정 - 저가
            #lastday_close = lastday[3]  #시고저종 설정 - 종가  #★★★
            #today_close = today[3]  #시고저종 설정 - 종가  #★★★
            targetday_TrMoney = targetday[5]  #시고저종 설정 - 거래대금  #★★★
            #lastday_vol = lastday[4]  #시고저종 설정 - 거래량  #★★★
            #target_price = today_open + (lastday_high - lastday_low) * 0.2   # 0.3
            return targetday_TrMoney
        except Exception as ex:
#             self.dbgout("`get_2daybef_TrMoney() -> exception! " + str(ex) + "`")
            return None

    def get_3daybef_TrMoney(self, code):
        """금일 시가를 반환한다."""
        try:
            time_now = datetime.now()  #현재 시간 추출
            str_today = time_now.strftime('%Y%m%d')
            ohlc = self.get_ohlc(code, 30)  #코드에 해당하는 30일치 ohlc값을 가져온다
            if len(ohlc) > 4:
                if str_today == str(ohlc.iloc[0].name):  #30일치중 첫번째값의 name이 오늘과 똑같으면,
                    today_open = ohlc.iloc[0].open  # today_open 설정
                    today = ohlc.iloc[0]  # today 설정
                    targetday = ohlc.iloc[3]  #어제값은 10일치 데이터중 두번째 값 설정
                else:
                    lastday = ohlc.iloc[0]  #만약10일치 데이터중 첫번째 값이 오늘이 아니라면(즉 당일조회가 아니라 밤12시 지나고 장시작전 조회할 경우), 10일치중 첫번째값이 어제데이터
                    today_open = lastday[3]  #금일 시가를 전일 종가로 설정
                    targetday = ohlc.iloc[3]  #어제값은 10일치 데이터중 두번째 값 설정
            else:
                targetday = ohlc.iloc[len(ohlc)-1]
            #lastday_high = lastday[1]  #시고저종 설정 - 고가
            #lastday_low = lastday[2]  #시고저종 설정 - 저가
            #lastday_close = lastday[3]  #시고저종 설정 - 종가  #★★★
            #today_close = today[3]  #시고저종 설정 - 종가  #★★★
            targetday_TrMoney = targetday[5]  #시고저종 설정 - 거래대금  #★★★
            #lastday_vol = lastday[4]  #시고저종 설정 - 거래량  #★★★
            #target_price = today_open + (lastday_high - lastday_low) * 0.2   # 0.3
            return targetday_TrMoney
        except Exception as ex:
#             self.dbgout("`get_3daybef_TrMoney() -> exception! " + str(ex) + "`")
            return None

    def get_4daybef_TrMoney(self, code):
        """금일 시가를 반환한다."""
        try:
            time_now = datetime.now()  #현재 시간 추출
            str_today = time_now.strftime('%Y%m%d')
            ohlc = self.get_ohlc(code, 30)  #코드에 해당하는 30일치 ohlc값을 가져온다
            if len(ohlc) > 5:
                if str_today == str(ohlc.iloc[0].name):  #30일치중 첫번째값의 name이 오늘과 똑같으면,
                    today_open = ohlc.iloc[0].open  # today_open 설정
                    today = ohlc.iloc[0]  # today 설정
                    targetday = ohlc.iloc[4]  #어제값은 10일치 데이터중 두번째 값 설정
                else:
                    lastday = ohlc.iloc[0]  #만약10일치 데이터중 첫번째 값이 오늘이 아니라면(즉 당일조회가 아니라 밤12시 지나고 장시작전 조회할 경우), 10일치중 첫번째값이 어제데이터
                    today_open = lastday[3]  #금일 시가를 전일 종가로 설정
                    targetday = ohlc.iloc[4]  #어제값은 10일치 데이터중 두번째 값 설정
            else:
                targetday = ohlc.iloc[len(ohlc)-1]
            #lastday_high = lastday[1]  #시고저종 설정 - 고가
            #lastday_low = lastday[2]  #시고저종 설정 - 저가
            #lastday_close = lastday[3]  #시고저종 설정 - 종가  #★★★
            #today_close = today[3]  #시고저종 설정 - 종가  #★★★
            targetday_TrMoney = targetday[5]  #시고저종 설정 - 거래대금  #★★★
            #lastday_vol = lastday[4]  #시고저종 설정 - 거래량  #★★★
            #target_price = today_open + (lastday_high - lastday_low) * 0.2   # 0.3
            return targetday_TrMoney
        except Exception as ex:
#             self.dbgout("`get_4daybef_TrMoney() -> exception! " + str(ex) + "`")
            return None











    def MktCapital(self, code):  # 시가총액구하는함수!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return MarketCapital1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_epsroe(self, code):  # "EPS x ROE = 적정주가" 로 대략생각가능!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return round(MarketStockprice1,1)  # MarketCapital1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_eps(self, code):  # "EPS x ROE = 적정주가" 로 대략생각가능!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return round(MktCpteye_eps1,1)  # MarketCapital1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_roe(self, code):  # "EPS x ROE = 적정주가" 로 대략생각가능!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return round(MktCpteye_roe1,1)  # MarketCapital1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_roe2(self, code):  # "EPS x ROE = 적정주가" 로 대략생각가능!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return round(MktCpteye_roe2,1)  # MarketCapital1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_pbr(self, code):  # "EPS x ROE = 적정주가" 로 대략생각가능!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return round(MarketPBR1,1)  # MarketCapital1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_per(self, code):  # "EPS x ROE = 적정주가" 로 대략생각가능!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return round(MktCpteye_per1,1)  # MarketCapital1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_debt(self, code):  # "EPS x ROE = 적정주가" 로 대략생각가능!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return round(MktCpteye_debt1,1)  # MarketCapital1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_moneyrate(self, code):  # "EPS x ROE = 적정주가" 로 대략생각가능!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return round(MktCpteye_moneyrate1,1)  # MarketCapital1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_bps(self, code):  # "EPS x ROE = 적정주가" 로 대략생각가능!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return MktCpteye_bps1  # MarketCapital1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_initialmoney(self, code):  # "EPS x ROE = 적정주가" 로 대략생각가능!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return MktCpteye_money1  # MarketCapital1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_stocknum(self, code):  # 시가총액구하는함수!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return MktCpteye_mktcpt1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_salesriserate1(self, code):  # 시가총액구하는함수!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return MktCpteye_salesriserate1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_salesriserate2(self, code):  # 시가총액구하는함수!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return MktCpteye_salesriserate2
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_prftriserate1(self, code):  # 시가총액구하는함수!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return MktCpteye_prftriserate1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_prftriserate2(self, code):  # 시가총액구하는함수!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return MktCpteye_prftriserate2
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_salesprft1(self, code):  # 시가총액구하는함수!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return MktCpteye_salesprft1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_salesprft2(self, code):  # 시가총액구하는함수!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return MktCpteye_salesprft2
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_salesprftrise1(self, code):  # 시가총액구하는함수!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return MktCpteye_salesprftrise1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_salesprftrise2(self, code):  # 시가총액구하는함수!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return MktCpteye_salesprftrise2
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)

    def MktCapital_beta1(self, code):  # 시가총액구하는함수!!
        MktCpteye1 = []
        # MktCapital_CodeList1 = []
        # MktCapital_NameList1 = []
        self.MarkCptlEye.SetInputValue(0, [0,4,17,20,23,67,70,71,75,76,77,78,80,89,90,92,97,98,100,105,107,150])  # 0 종목코드, 4 현재가, 17 종목명, 20 총상장주식수(단위:주), 23 전일종가, 67 PER, 70 EPS, 71 자본금(백만), 75 부채비율, 76 유보율(%), 77 ROE자기자본순이익률, 
        # 78 매출액증가율, 80 순이익증가율, 89 BPS 주당순자산, 90 영업이익증가율, 92 매출대비영업이익률, 97 분기매출액증가율, 98 분기영업이익증가율, 100 분기순이익증가율, 105 분기매출대비영업이익률, 107 분기ROE, 150 베타계수
        self.MarkCptlEye.SetInputValue(1, code)
        self.MarkCptlEye.BlockRequest()
        markcpteyelist1 = self.MarkCptlEye.GetHeaderValue(2)
        for i in range(markcpteyelist1):
            MktCpteye_code1 = self.MarkCptlEye.GetDataValue(0, i)  # 0 종목코드
            MktCpteye_todayprice1 = self.MarkCptlEye.GetDataValue(1, i)  # 4 현재가
            MktCpteye_codename1 = self.MarkCptlEye.GetDataValue(2, i)  # 17 종목명
            MktCpteye_mktcpt1 = self.MarkCptlEye.GetDataValue(3, i)  # 20 총상장주식수(단위: 주)
            MktCpteye_curprc1 = self.MarkCptlEye.GetDataValue(4, i)  # 23 전일종가
            MktCpteye_per1 = self.MarkCptlEye.GetDataValue(5, i)  # 67 PER (보통 10~30정도가 적정)
            MktCpteye_eps1 = self.MarkCptlEye.GetDataValue(6, i)  # 70 EPS -> "PER x EPS = 적정주가" 라고 대략생각가능!!  -> but, EPS x ROE 가 보다더정확한 적정주가계산공식인듯!!
            MktCpteye_money1 = self.MarkCptlEye.GetDataValue(7, i)  # 71 자본금 (백만)
            MktCpteye_debt1 = self.MarkCptlEye.GetDataValue(8, i)  # 75 부채비율
            MktCpteye_moneyrate1 = self.MarkCptlEye.GetDataValue(9, i)  # 76 유보율 (%)
            MktCpteye_roe1 = self.MarkCptlEye.GetDataValue(10, i)  # 77 "ROE 자기자본순이익률 = PBR/PER"
            MktCpteye_salesriserate1 = self.MarkCptlEye.GetDataValue(11, i)  # 78 매출액증가율
            MktCpteye_prftriserate1 = self.MarkCptlEye.GetDataValue(12, i)  # 80 순이익증가율
            MktCpteye_bps1 = self.MarkCptlEye.GetDataValue(13, i)  # 89 BPS 주당순자산
            MktCpteye_salesprftrise1 = self.MarkCptlEye.GetDataValue(14, i)  # 90 영업이익증가율
            MktCpteye_salesprft1 = self.MarkCptlEye.GetDataValue(15, i)  # 92 매출대비영업이익률
            MktCpteye_salesriserate2 = self.MarkCptlEye.GetDataValue(16, i)  # 97 분기매출액증가율
            MktCpteye_salesprftrise2 = self.MarkCptlEye.GetDataValue(17, i)  # 98 분기영업이익증가율
            MktCpteye_prftriserate2 = self.MarkCptlEye.GetDataValue(18, i)  # 100 분기순이익증가율
            MktCpteye_salesprft2 = self.MarkCptlEye.GetDataValue(19, i)  # 105 분기매출대비영업이익률
            MktCpteye_roe2 = self.MarkCptlEye.GetDataValue(20, i)  # 107 분기ROE
            MktCpteye_beta1 = self.MarkCptlEye.GetDataValue(21, i)  # 150 베타계수
            MktCpteye1.append( {'MktCptEyecode1': MktCpteye_code1, 'MktCptEyetodayprice1': MktCpteye_todayprice1, 'MktCptEyename1': MktCpteye_codename1, 'MktCptEyecpt1': MktCpteye_mktcpt1, 'MktCptEyeprc1': MktCpteye_curprc1, \
                'MktCptEyeper1': MktCpteye_per1, 'MktCptEyeeps1': MktCpteye_eps1, 'MktCptEyemoney1': MktCpteye_money1, 'MktCptEyedebt1': MktCpteye_debt1, 'MktCptEyemoneyrate1': MktCpteye_moneyrate1, \
                'MktCptEyeroe1': MktCpteye_roe1, 'MktCptEyesalesriserate1': MktCpteye_salesriserate1, 'MktCptEyeprftriserate1': MktCpteye_prftriserate1, 'MktCptEyesalesprft1': MktCpteye_salesprft1, 'MktCptEyesalesprftrise1': MktCpteye_salesprftrise1, 'MktCptEyebps1': MktCpteye_bps1, \
                'MktCptEyeroe2': MktCpteye_roe2, 'MktCptEyesalesriserate2': MktCpteye_salesriserate2, 'MktCptEyeprftriserate2': MktCpteye_prftriserate2, 'MktCptEyesalesprft2': MktCpteye_salesprft2, 'MktCptEyesalesprftrise2': MktCpteye_salesprftrise2, 'MktCptEyebeta1': MktCpteye_beta1 } )
            MarketCapital1 = MktCpteye_mktcpt1 * MktCpteye_curprc1
            MarketStockprice1 = MktCpteye_eps1 * MktCpteye_roe1  # -> "EPS x ROE = 적정주가" 라고 대략생각가능!!
            MarketPBR1 = ( ( (MktCpteye_roe1 * MktCpteye_per1)/100 ) + ( MktCpteye_todayprice1/MktCpteye_bps1 ) )/2  # "PBR = ROE x PER" 혹은 "PBR = 주가/BPS" 로 계산함!!
            if self.instCpCodeMgr.IsBigListingStock(code) :
                MarketCapital1 *= 1000  #상장주식수 20억이상시, 시가총액이 1000배적게 나오는문제가 있어서, 이럴때에만 *1000 을 곱하도록되어있음!!
            return MktCpteye_beta1
            # if ( MarketCapital1 ) > 1 :
            #     self.dbgout( str(MktCpteye_codename1) + ', 시가총액 ' + str( round( (MarketCapital1)/1000000000000, 4) ) + '조원 ▨' )
            #     MktCapital_CodeList1.append( MktCpteye_code1 )
            #     MktCapital_NameList1.append( MktCpteye_codename1 )
            #     MarketCptEyelist1 = pd.DataFrame( [ MktCapital_NameList1, MktCapital_CodeList1 ] )
            #     MarketCptEyelist1.transpose().to_excel("_99MktCpt_KSP.xlsx", index=False, header=False)





    # if __name__ == "__main__":
