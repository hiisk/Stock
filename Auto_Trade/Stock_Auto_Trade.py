import os, sys, ctypes
import win32com.client
import pandas as pd
from datetime import datetime
from slacker import Slacker
import time, calendar
from urllib.request import urlopen
import numpy as np

import requests
def post_message(token, channel, text):
    response = requests.post("https://slack.com/api/chat.postMessage",
        headers={"Authorization": "Bearer "+token},
        data={"channel": channel,"text": text}
    )
 
myToken = "your Token"
def dbgout(message):
    """인자로 받은 문자열을 파이썬 셸과 슬랙으로 동시에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
    post_message(myToken,"#stock",strbuf)

def printlog(message, *args):
    """인자로 받은 문자열을 파이썬 셸에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message, *args)
 
# 크레온 플러스 공통 OBJECT
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')  

def check_creon_system():
    """크레온 플러스 시스템 연결 상태를 점검한다."""
    # 관리자 권한으로 프로세스 실행 여부
    if not ctypes.windll.shell32.IsUserAnAdmin():
        printlog('check_creon_system() : admin user -> FAILED')
        return False
 
    # 연결 여부 체크
    if (cpStatus.IsConnect == 0):
        printlog('check_creon_system() : connect to server -> FAILED')
        return False
 
    # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    if (cpTradeUtil.TradeInit(0) != 0):
        printlog('check_creon_system() : init trade -> FAILED')
        return False
    return True

def get_current_price(code):
    """인자로 받은 종목의 현재가, 매수호가, 매도호가를 반환한다."""
    cpStock.SetInputValue(0, code)  # 종목코드에 대한 가격 정보
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)   # 현재가
    item['ask'] =  cpStock.GetHeaderValue(16)        # 매수호가
    item['bid'] =  cpStock.GetHeaderValue(17)        # 매도호가    
    return item['cur_price'], item['ask'], item['bid']

def get_ohlc(code, qty):
    """인자로 받은 종목의 OHLC 가격 정보를 qty 개수만큼 반환한다."""
    cpOhlc.SetInputValue(0, code)           # 종목코드
    cpOhlc.SetInputValue(1, ord('2'))        # 1:기간, 2:개수
    cpOhlc.SetInputValue(4, qty)             # 요청개수
    cpOhlc.SetInputValue(5, [0, 2, 3, 4, 5]) # 0:날짜, 2~5:OHLC
    cpOhlc.SetInputValue(6, ord('D'))        # D:일단위
    cpOhlc.SetInputValue(9, ord('1'))        # 0:무수정주가, 1:수정주가
    cpOhlc.BlockRequest()
    count = cpOhlc.GetHeaderValue(3)   # 3:수신개수
    columns = ['open', 'high', 'low', 'close']
    index = []
    rows = []
    for i in range(count): 
        index.append(cpOhlc.GetDataValue(0, i)) 
        rows.append([cpOhlc.GetDataValue(1, i), cpOhlc.GetDataValue(2, i),
            cpOhlc.GetDataValue(3, i), cpOhlc.GetDataValue(4, i)]) 
    df = pd.DataFrame(rows, columns=columns, index=index) 
    return df

def get_stock_balance(code):
    """인자로 받은 종목의 종목명과 수량을 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpBalance.SetInputValue(0, acc)         # 계좌번호
    cpBalance.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpBalance.SetInputValue(2, 50)          # 요청 건수(최대 50)
    cpBalance.BlockRequest()     
    if code == 'ALL':
        dbgout('계좌명: ' + str(cpBalance.GetHeaderValue(0)))
        dbgout('결제잔고수량 : ' + str(cpBalance.GetHeaderValue(1)))
        dbgout('평가금액: ' + str(cpBalance.GetHeaderValue(3)))
        dbgout('평가손익: ' + str(cpBalance.GetHeaderValue(4)))
        dbgout('종목수: ' + str(cpBalance.GetHeaderValue(7)))
    stocks = []
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)  # 종목코드
        stock_name = cpBalance.GetDataValue(0, i)   # 종목명
        stock_qty = cpBalance.GetDataValue(15, i)   # 수량
        if code == 'ALL':
            dbgout(str(i+1) + ' ' + stock_code + '(' + stock_name + ')' 
                + ':' + str(stock_qty))
            stocks.append({'code': stock_code, 'name': stock_name, 
                'qty': stock_qty})
        if stock_code == code:  
            return stock_name, stock_qty
    if code == 'ALL':
        return stocks
    else:
        stock_name = cpCodeMgr.CodeToName(code)
        return stock_name, 0

def get_current_cash():
    """증거금 100% 주문 가능 금액을 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]    # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpCash.SetInputValue(0, acc)              # 계좌번호
    cpCash.SetInputValue(1, accFlag[0])      # 상품구분 - 주식 상품 중 첫번째
    cpCash.BlockRequest() 
    return cpCash.GetHeaderValue(9) # 증거금 100% 주문 가능 금액

def get_target_price(code):
    """매수 목표가를 반환한다."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc_2weeks = get_ohlc(code, 10) # 2주
        ohlc_3days = ohlc_2weeks.iloc[:3,:4] # 3일
        
        if str_today == str(ohlc_2weeks.iloc[0].name):
            today_open = ohlc_2weeks.iloc[0].open 
            lastday = ohlc_2weeks.iloc[1]
        else:
            lastday = ohlc_2weeks.iloc[0]                                      
            today_open = lastday[3]
        
        lastday_high = lastday[1]
        lastday_low = lastday[2]
        lastday_open = lastday[0]
        lastday_close = lastday[3]

        global symbol_list_rate

        k = 0
        target_tmp = 1

        for i in np.arange(0.1, 1.0, 0.01): # 백테스팅을 통한 최적 K값 도출
            ohlc_2weeks['range'] = (ohlc_2weeks['high'] - ohlc_2weeks['low']) * i #2주 백테스팅
            ohlc_2weeks['target'] = ohlc_2weeks['open'] + ohlc_2weeks['range'].shift(1)

            ohlc_3days['range'] = (ohlc_3days['high'] - ohlc_3days['low']) * i #3일 백테스팅
            ohlc_3days['target'] = ohlc_3days['open'] + ohlc_3days['range'].shift(1)

            target_2weeks = np.where(ohlc_2weeks['high'] > ohlc_2weeks['target'], ohlc_2weeks['close'] / ohlc_2weeks['target'],1)
            target_3days = np.where(ohlc_3days['high'] > ohlc_3days['target'], ohlc_3days['close'] / ohlc_3days['target'],1)
            
            if target_tmp < target_2weeks.cumprod()[-2]  and 1.003 < target_3days.cumprod()[-2]: #3일 백테스팅이 1.003보다 클경우
                target_tmp = target_2weeks.cumprod()[-2]
                k = i
        
        if k == 0 or target_tmp < 1.003: # 2주 백테스팅의 수익률이 낮다면 제거
            delete_list.append(code)

        else: 
            if target_tmp > 1.01:
                target_tmp = (target_tmp-1)/3 + 1
                print(code, k, target_tmp)
        symbol_list_rate[code] = target_tmp

        target_price = today_open + int((lastday_high - lastday_low) * k) # k값이 낮게 설정되있는 경향을 자주보임(0.1일 경우가 많음)
        
        return target_price

    except Exception as ex:
        dbgout("`get_target_price() -> exception! " + str(ex) + "`")
        return None
    
def get_movingaverage(code, window):
    """인자로 받은 종목에 대한 이동평균가격을 반환한다."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 20)
        if str_today == str(ohlc.iloc[0].name):
            lastday = ohlc.iloc[1].name
        else:
            lastday = ohlc.iloc[0].name
        closes = ohlc['close'].sort_index()         
        ma = closes.rolling(window=window).mean()
        return ma.loc[lastday]
    except Exception as ex:
        dbgout('get_movingavrg(' + str(window) + ') -> exception! ' + str(ex))
        return None    

def stock_trade(code):
    """인자로 받은 종목을 시장가 IOC 조건으로 매수한다."""
    try:
        global bought_list      # 함수 내에서 값 변경을 하기 위해 global로 지정
        global sold_list      # 함수 내에서 값 변경을 하기 위해 global로 지정  

        # if code in bought_list: # 매수 완료 종목이면 더 이상 안 사도록 함수 종료
        #     #printlog('code:', code, 'in', bought_list)
        #     return False
        time_now = datetime.now()
        current_price, ask_price, bid_price = get_current_price(code) 
        
        # target_price = get_target_price(code)    # 매수 목표가
        if code not in bought_list:
            ma5_price = get_movingaverage(code, 5)   # 5일 이동평균가
            ma10_price = get_movingaverage(code, 10) # 10일 이동평균가
        buy_qty = 0        # 매수할 수량 초기화
        if ask_price > 0:  # 매수호가가 존재하면   
            buy_qty = buy_amount // ask_price  
        stock_name, stock_qty = get_stock_balance(code)  # 종목명과 보유수량 조회

        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체,1:주식,2:선물/옵션    
        current_cash = int(get_current_cash()) # 증거금 100% 주문 가능 금액

        if ((code not in bought_list) and current_price > symbol_list_value[code] and current_price > ma5_price and current_price > ma10_price
            and current_cash > total_cash * buy_percent and (current_price < symbol_list_value[code] * 1.002)) :       

            cpOrder.SetInputValue(0, "2")        # 2: 매수
            cpOrder.SetInputValue(1, acc)        # 계좌번호
            cpOrder.SetInputValue(2, accFlag[0]) # 상품구분 - 주식 상품 중 첫번째
            cpOrder.SetInputValue(3, code)       # 종목코드
            cpOrder.SetInputValue(4, buy_qty)    # 매수할 수량
            cpOrder.SetInputValue(7, "1")        # 주문조건 0:기본, 1:IOC, 2:FOK
            cpOrder.SetInputValue(8, "03")       # 주문호가 1:보통, 3:시장가
                                                 # 5:조건부, 12:최유리, 13:최우선 
            # 매수 주문 요청
            dbgout('시장가 IOC 조건 ' + '\n' + str(stock_name) + '\t' + str(code) + '\n' + '매수 완료')
            bought_list.append(code)

            ret = cpOrder.BlockRequest() 
            if ret == 4:
                remain_time = cpStatus.LimitRequestRemainTime
                printlog('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time/1000)
                time.sleep(remain_time/1000) 
                return False
        
        #반 매도
        elif (code not in sold_list) and (code in bought_list) and stock_qty != 0 and current_price >= symbol_list_value[code] * symbol_list_rate[code] :
            cpOrder.SetInputValue(0, "1")         # 1:매도, 2:매수
            cpOrder.SetInputValue(1, acc)         # 계좌번호
            cpOrder.SetInputValue(2, accFlag[0])  # 주식상품 중 첫번째
            cpOrder.SetInputValue(3, code)   # 종목코드
            cpOrder.SetInputValue(4, round(stock_qty/2))    # 매도수량
            cpOrder.SetInputValue(7, "1")   # 조건 0:기본, 1:IOC, 2:FOK
            cpOrder.SetInputValue(8, "12")  # 호가 12:최유리, 13:최우선 
            # 시장가 IOC 매도 주문 요청

            dbgout('최유리 IOC 조건 ' + '\n' + str(stock_name) + '\t' + str(code) + '\n' + '반 매도 완료')
            sold_list.append(code)

            ret = cpOrder.BlockRequest() 
            if ret == 4:
                remain_time = cpStatus.LimitRequestRemainTime
                printlog('주의: 연속 주문 제한, 대기시간:', remain_time/1000)
        
        #반매도후 나머지 매도
        elif (code in sold_list) and (code in bought_list) and stock_qty != 0 and current_price >= symbol_list_value[code] * ((symbol_list_rate[code]-1)*3+1) :
            cpOrder.SetInputValue(0, "1")         # 1:매도, 2:매수
            cpOrder.SetInputValue(1, acc)         # 계좌번호
            cpOrder.SetInputValue(2, accFlag[0])  # 주식상품 중 첫번째
            cpOrder.SetInputValue(3, code)   # 종목코드
            cpOrder.SetInputValue(4, stock_qty)    # 매도수량
            cpOrder.SetInputValue(7, "1")   # 조건 0:기본, 1:IOC, 2:FOK
            cpOrder.SetInputValue(8, "03")  # 호가 12:최유리, 13:최우선 
            # 시장가 IOC 매도 주문 요청
            dbgout('시장가 IOC 조건 ' + '\n' + str(stock_name) + '\t' + str(code) + '\n' + '전체 매도 완료')
            
            symbol_list.remove(code)
            bought_list.remove(code)
            sold_list.remove(code)

            ret = cpOrder.BlockRequest()
            if ret == 4:
                remain_time = cpStatus.LimitRequestRemainTime
                printlog('주의: 연속 주문 제한, 대기시간:', remain_time/1000)   

    except Exception as ex:
        dbgout("`stock_trade("+ str(code) + ") -> exception! " + str(ex) + "`")

def sell_all():
    """보유한 모든 종목을 시장가 IOC 조건으로 매도한다."""

    try:
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]       # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션   
        while True:    
            stocks = get_stock_balance('ALL') 
            total_qty = 0 
            for s in stocks:
                total_qty += s['qty'] 
            if total_qty == 0:
                return True
            for s in stocks:
                if s['qty'] != 0:                  
                    cpOrder.SetInputValue(0, "1")         # 1:매도, 2:매수
                    cpOrder.SetInputValue(1, acc)         # 계좌번호
                    cpOrder.SetInputValue(2, accFlag[0])  # 주식상품 중 첫번째
                    cpOrder.SetInputValue(3, s['code'])   # 종목코드
                    cpOrder.SetInputValue(4, s['qty'])    # 매도수량
                    cpOrder.SetInputValue(7, "1")   # 조건 0:기본, 1:IOC, 2:FOK
                    cpOrder.SetInputValue(8, "03")  # 호가 12:최유리, 13:최우선, 3:시장가
                    # 시장가 IOC 조건 주문 요청
                    ret = cpOrder.BlockRequest()
                    dbgout('시장가 IOC 조건 ' + '\n' + s['name'] + '\t' + s['code'] + '\n' + '매도 완료')
                    if ret == 4:
                        remain_time = cpStatus.LimitRequestRemainTime
                        printlog('주의: 연속 주문 제한, 대기시간:', remain_time/1000)
                time.sleep(1)
            time.sleep(10)
    except Exception as ex:
        dbgout("sell_all() -> exception! " + str(ex))

if __name__ == '__main__': 
    try:
        #전체 코드
        objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
        codeList = objCodeMgr.GetStockListByMarket(1)  # 거래소
        codeList2 = objCodeMgr.GetStockListByMarket(2)  # 코스닥
        allCode = codeList + codeList2
        ETFList = []
        for code in allCode:
            stockKind = objCodeMgr.GetStockSectionKind(code)
    
            if stockKind == 10 or stockKind == 12 :
                ETFList.append(code)
        symbol_list = []
        symbol_list_value = {}
        symbol_list_rate = {}

        delete_list = []

        buy_percent = 0.1
        total_cash = int(get_current_cash())   # 100% 증거금 주문 가능 금액 조회
        buy_amount = int(total_cash * buy_percent)  # 종목별 주문 금액 계산

        objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        
        # 차트 객체 구하기
        for i in range(0, len(ETFList)):
            objStockChart.SetInputValue(0, ETFList[i])   #종목 코드 - 삼성전자
            objStockChart.SetInputValue(1, ord('2')) # 개수로 조회
            objStockChart.SetInputValue(4, 100) # 최근 100일 치
            objStockChart.SetInputValue(5, [0,2,3,4,5, 8]) #날짜,시가,고가,저가,종가,거래량
            objStockChart.SetInputValue(6, ord('D')) # '차트 주가 - 일간 차트 요청
            objStockChart.SetInputValue(9, ord('1')) # 수정주가 사용
            objStockChart.BlockRequest()
            vol = objStockChart.GetDataValue(5, 1) #전날 거래량 = 1
            
            if vol > 10000 :
                target_price = get_target_price(ETFList[i])
                if target_price < buy_amount:  #목표가가 종목별 주문 금액보다 클 때 삭제
                    symbol_list.append(ETFList[i])
                    symbol_list_value[ETFList[i]] = target_price
        
            time.sleep(0.25)
        
        dbgout("삭제전 종목 개수: "+ str(len(symbol_list)) +", "+ str(len(symbol_list_value))+", "+ str(len(symbol_list_rate)))
        
        if len(delete_list) > 0:
            for d in delete_list: # 수익률 낮은 종목 삭제
                if d in symbol_list and d in symbol_list_value:
                    symbol_list.remove(d)
                    del(symbol_list_value[d])
                    del(symbol_list_rate[d])

        bought_list = []     # 매수 완료된 종목 리스트
        sold_list = []      # 반 매도 완료된 종목 리스트

        dbgout('삭제후 종목 개수: '+ str(len(symbol_list)) +", "+ str(len(symbol_list_value)) +", "+ str(len(symbol_list_rate)))
        dbgout('100% 증거금 주문 가능 금액: ' + str(total_cash))
        dbgout('종목별 주문 비율: ' + str(buy_percent))
        dbgout('종목별 주문 금액: ' + str(buy_amount))
        printlog('시작 시간: ', datetime.now().strftime('%m/%d %H:%M:%S'))
        printlog('check_creon_system(): ', check_creon_system())  # 크레온 접속 점검
        # soldout = False

        while True:
            t_now = datetime.now()
            t_9 = t_now.replace(hour=9, minute=0, second=0, microsecond=0)
            t_start = t_now.replace(hour=9, minute=0, second=0, microsecond=0)
            t_sell = t_now.replace(hour=15, minute=10, second=0, microsecond=0)
            t_exit = t_now.replace(hour=15, minute=20, second=0,microsecond=0)
            today = datetime.today().weekday()
            if today == 5 or today == 6:  # 토요일이나 일요일이면 자동 종료
                printlog('Today is', 'Saturday.' if today == 5 else 'Sunday.')
                sys.exit(0)
            # if t_9 < t_now < t_start and soldout == False:
            #     soldout = True
            #     sell_all()
            if t_start < t_now < t_sell :  # AM 09:00 ~ PM 03:10 : 매수
                current_cash = int(get_current_cash()) # 증거금 100% 주문 가능 금액
                if current_cash > total_cash * buy_percent:
                    for sym in symbol_list:
                        stock_trade(sym)
                        time.sleep(1)
                elif current_cash < total_cash * buy_percent: # 증거금이 부족할 때 매도 최적화      
                    for bou in bought_list:
                        stock_trade(bou)
                        time.sleep(1)
                if t_now.minute == 0 and 0 <= t_now.second <= 5: #한 시간마다 한번씩 계좌 정보 조회
                    get_stock_balance('ALL')
                    time.sleep(5)
            if t_sell < t_now < t_exit:  # PM 03:10 ~ PM 03:20 : 일괄 매도
                if sell_all() == True:
                    dbgout('매도끝 프로그램 종료')
                    sys.exit(0)      
            if t_exit < t_now:  # PM 03:20 ~ :프로그램 종료
                dbgout('프로그램 종료')
                sys.exit(0)
            time.sleep(3)
    except Exception as ex:
        dbgout('`main -> exception! ' + str(ex) + '`')