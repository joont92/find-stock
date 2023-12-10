import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes

import os
import time
from pywinauto import application

################################################
# PLUS 공통 OBJECT
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')


# 로그인 여부 체크
def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False
 
    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False
 
    '''
    # 주문 관련 초기화
    if (g_objCpTrade.TradeInit(0) != 0):
        print("주문 초기화 실패")
        return False
    '''
    return True

# 자동 로그인
def connect(reconnect=True):
    # 재연결이라면 기존 연결을 강제로 kill
    if reconnect:
        try:
            os.system('taskkill /IM ncStarter* /F /T')
            os.system('taskkill /IM CpStart* /F /T')
            os.system('taskkill /IM DibServer* /F /T')
            os.system('wmic process where "name like \'%ncStarter%\'" call terminate')
            os.system('wmic process where "name like \'%CpStart%\'" call terminate')
            os.system('wmic process where "name like \'%DibServer%\'" call terminate')
        except:
            pass

    CpCybos = win32com.client.Dispatch("CpUtil.CpCybos")

    if CpCybos.IsConnect:
        print('already connected.')

    else:
        app = application.Application()
        app.start(
            'C:\Daishin\Starter\\ncStarter.exe /prj:cp /id:{id} /pwd:{pwd} /pwdcert:{pwdcert} /autostart'.format(
                id='******', pwd='******', pwdcert='******')
        )
        # 연결 될때까지 무한루프
        while True:
            if CpCybos.IsConnect:
                break
            time.sleep(1)

        print('connected.')
    return CpCybos

# 특징주 포착
# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관
        self.diccode = {
            10: '외국계증권사창구첫매수',
            11: '외국계증권사창구첫매도',
            12: '외국인순매수',
            13: '외국인순매도',
            21: '전일거래량갱신',
            22: '최근5일거래량최고갱신',
            23: '최근5일매물대돌파',
            24: '최근60일매물대돌파',
            28: '최근5일첫상한가',
            29: '최근5일신고가갱신',
            30: '최근5일신저가갱신',
            31: '상한가직전',
            32: '하한가직전',
            41: '주가 5MA 상향돌파',
            42: '주가 5MA 하향돌파',
            43: '거래량 5MA 상향돌파',
            44: '주가데드크로스(5MA < 20MA)',
            45: '주가골든크로스(5MA > 20MA)',
            46: 'MACD 매수-Signal(9) 상향돌파',
            47: 'MACD 매도-Signal(9) 하향돌파',
            48: 'CCI 매수-기준선(-100) 상향돌파',
            49: 'CCI 매도-기준선(100) 하향돌파',
            50: 'Stochastic(10,5,5)매수- 기준선상향돌파',
            51: 'Stochastic(10,5,5)매도- 기준선하향돌파',
            52: 'Stochastic(10,5,5)매수- %K%D 교차',
            53: 'Stochastic(10,5,5)매도- %K%D 교차',
            54: 'Sonar 매수-Signal(9) 상향돌파',
            55: 'Sonar 매도-Signal(9) 하향돌파',
            56: 'Momentum 매수-기준선(100) 상향돌파',
            57: 'Momentum 매도-기준선(100) 하향돌파',
            58: 'RSI(14) 매수-Signal(9) 상향돌파',
            59: 'RSI(14) 매도-Signal(9) 하향돌파',
            60: 'Volume Oscillator 매수-Signal(9) 상향돌파',
            61: 'Volume Oscillator 매도-Signal(9) 하향돌파',
            62: 'Price roc 매수-Signal(9) 상향돌파',
            63: 'Price roc 매도-Signal(9) 하향돌파',
            64: '일목균형표매수-전환선 > 기준선상향교차',
            65: '일목균형표매도-전환선 < 기준선하향교차',
            66: '일목균형표매수-주가가선행스팬상향돌파',
            67: '일목균형표매도-주가가선행스팬하향돌파',
            68: '삼선전환도-양전환',
            69: '삼선전환도-음전환',
            70: '캔들패턴-상승반전형',
            71: '캔들패턴-하락반전형',
            81: '단기급락후 5MA 상향돌파',
            82: '주가이동평균밀집-5%이내',
            83: '눌림목재상승-20MA 지지'
        }
 
    def OnReceived(self):
        print(self.name)
        # 실시간 처리 - marketwatch : 특이 신호(차트, 외국인 순매수 등)
        if self.name == 'marketwatch':
            code = self.client.GetHeaderValue(0)
            name = g_objCodeMgr.CodeToName(code)
            cnt = self.client.GetHeaderValue(2)
 
            for i in range(cnt):
                item = {}
                newcancel = ''
                time = self.client.GetDataValue(0, i)
                h,m = divmod(time, 100)
                item['시간'] = '%02d:%02d' % (h,m)
                update = self.client.GetDataValue(1, i)
                item['코드'] = code
                item['종목명'] = name
                cate = self.client.GetDataValue(2, i)
                if (update == ord('c')):
                    newcancel =  '[취소]'
                if cate in self.diccode:
                    item['특이사항'] = newcancel + self.diccode[cate]
                else:
                    item['특이사항'] = newcancel + ''
 
                self.caller.listWatchData.insert(0,item)
                print(item)
 
        # 실시간 처리 - marketnews : 뉴스 및 공시 정보
        elif self.name == 'marketnews':
            item = {}
            update = self.client.GetHeaderValue(0)
            cont = ''
            if update == ord('D') :
                cont = '[삭제]'
            code = item['코드'] = self.client.GetHeaderValue(1)
            time = self.client.GetHeaderValue(2)
            h, m = divmod(time, 100)
            item['시간'] = '%02d:%02d' % (h, m)
            item['종목명'] = name = g_objCodeMgr.CodeToName(code)
            cate = self.client.GetHeaderValue(4)
            item['특이사항'] = cont + self.client.GetHeaderValue(5)
            print(item)
            self.caller.listWatchData.insert(0, item)
 
class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False
 
    def Subscribe(self, var, caller):
        if self.bIsSB:
            self.Unsubscribe()
 
        if (len(var) > 0):
            self.obj.SetInputValue(0, var)
 
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, caller)
        self.obj.Subscribe()
        self.bIsSB = True
 
    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
        self.bIsSB = False

class CpRpMarketWatch:
    def __init__(self):
        self.objStockMst = win32com.client.Dispatch('CpSysDib.CpMarketWatch') # 특징주 포착
        return
 
    # self.objMarketWatch.Request('*', self)
    def Request(self, listWatchData): 
        self.objStockMst.SetInputValue(0, '*') # *: 전종목
        self.objStockMst.SetInputValue(1, '1') # 1: 종목 뉴스 2: 공시정보 10: 외국계 창구첫매수, 11:첫매도 12 외국인 순매수 13 순매도
        self.objStockMst.SetInputValue(2, 0) # 시작 시간: 0 처음부터
 
        flag = True
        while flag:
            ret = self.objStockMst.BlockRequest()
            if self.objStockMst.GetDibStatus() != 0:
                print('통신상태', self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
                return False
     
            #if self.objStockMst.Continue == 0:
            flag = False
            cnt = self.objStockMst.GetHeaderValue(2)  # 수신 개수
            for i in range(cnt):
                # item = {}
     
                # time  = self.objStockMst.GetDataValue(0, i)
                # h, m = divmod(time, 100)
                # item['시간'] = '%02d:%02d' % (h, m)
                # item['코드'] = self.objStockMst.GetDataValue(1, i)
                # item['종목명'] = g_objCodeMgr.CodeToName(item['코드'])
                # cate = self.objStockMst.GetDataValue(3, i)
                # item['특이사항'] = self.objStockMst.GetDataValue(4, i)
                listWatchData.append(self.objStockMst.GetDataValue(1, i))

        return True

# 차트 기본 데이터 통신
class CpStockChart:
    def __init__(self):
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        self.objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        # CpIndexes.CpSeries : 차트 기본 데이터 관리 PLUS 객체
        self.objSeries = win32com.client.Dispatch("CpIndexes.CpSeries")
        self.result = []

    def Request(self, code, cnt):
    	#######################################################
        # 1. 일간 차트 데이터 요청
        self.objStockChart.SetInputValue(0, code)  # 종목 코드 -
        self.objStockChart.SetInputValue(1, ord('2'))  # 개수로 조회
        self.objStockChart.SetInputValue(4, cnt)  # 최근 {cnt}일치
        self.objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8, 9])  # 날짜,시가,고가,저가,종가,거래량,거래대금
        self.objStockChart.SetInputValue(6, ord('D'))  # '차트 주기 - 일간 차트 요청
        self.objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
        self.objStockChart.BlockRequest()

        rqStatus = self.objStockChart.GetDibStatus()
        rqRet = self.objStockChart.GetDibMsg1()
        # print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()

        #######################################################
        # 2. 일간 차트 데이터 ==> CpIndexes.CpSeries 로 변환
        len = self.objStockChart.GetHeaderValue(3)

        # print("날짜", "시가", "고가", "저가", "종가", "거래량")
        # print("==============================================-")
        for i in range(len):
            day = self.objStockChart.GetDataValue(0, len - i - 1)
            open = self.objStockChart.GetDataValue(1, len - i - 1)
            high = self.objStockChart.GetDataValue(2, len - i - 1)
            low = self.objStockChart.GetDataValue(3, len - i - 1)
            close = self.objStockChart.GetDataValue(4, len - i - 1)
            vol = self.objStockChart.GetDataValue(5, len - i - 1)
            amount = self.objStockChart.GetDataValue(6, len - i - 1)
            # print(day, open, high, low, close, vol)
            # objSeries.Add 종가, 시가, 고가, 저가, 거래량, 코멘트
            self.objSeries.Add(close, open, high, low, vol)
            self.result.append({"day": day, "open": open, "close": close, "high": high, "low": low, "vol": vol, "amount":  amount})
        # print("==============================================-")

        return

# 이동평균선
class CpIndex:
    def __init__(self):
        # CpIndexes.CpIndex : 지표 계산을 담당하는 PLrt.GetHeaderVaUS 객체
        self.objIndex = win32com.client.Dispatch("CpIndexes.CpIndex")


    # 주어진 지표의 이름(indexName) 으로 지표 계산 및 데이터 리턴
    def makeIndex(self, indexValue, objSeries):
        self.objIndex.series = objSeries
        self.objIndex.put_IndexKind('이동평균(라인1개)')  # 계산할 지표:
        # self.objIndex.put_IndexDefault(indexName)  # 지표 기본 변수 자동 세팅

        # 지표 데이터 계산 하기
        self.objIndex.Term1 = 240 # 240일선
        self.objIndex.Calculate()

        cntofIndex = self.objIndex.ItemCount
        # print("지표 개수:  ", cntofIndex)
        # 지표의 각 라인 이름은 HTS 차트의 각 지표 조건 참고
        for index in range(cntofIndex):
            cnt = self.objIndex.GetCount(index)
            for j in range(cnt) :
                value = self.objIndex.GetResult(index,j)
                indexValue.append(value)


# 시작
connect(False)

if InitPlusCheck() == False:
	exit()

# 특징주 뽑기
listWatchData = []
objMarketWatch = CpRpMarketWatch()
objMarketWatch.Request(listWatchData)
listWatchData = list(set(listWatchData))


for c in listWatchData:
    # 이평선 구하기
    cnt = 240
    objCpChart = CpStockChart()
    objCpChart.Request(c, cnt + 1) # 240일선 구하려면 +1일치 차트를 불러와야하는 듯

    objCpIndex = CpIndex()
    indexData = []
    objCpIndex.makeIndex(indexData, objCpChart.objSeries)
    
    # 오늘자
    todayChart = objCpChart.result
    open = todayChart[cnt]['open']
    close = todayChart[cnt]['close']
    high = todayChart[cnt]['high']
    low = todayChart[cnt]['low']
    vol = todayChart[cnt]['vol']
    amount = todayChart[cnt]['amount']
    # 오늘자 240일 이평선 가격
    movingAverage240 = indexData[cnt - 1]

    # 거래금액 필터?
    천억 = 100000000000
    if (amount >= todayChart[cnt - 1]['amount'] + 천억):
        print("종목:", c, ", 오늘 거래량: ", amount, ", 어제 거래량: ", todayChart[cnt - 1]['amount'], ", 오늘자 차트 데이터: ", todayChart[cnt])