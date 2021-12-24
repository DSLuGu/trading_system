# -*- coding:utf-8 -*-
# author: LuGu
# initial date: 15/12/2021
# last edited date: 24/12/2021
# description: creonAPI
# reference: https://github.com/gyusu/Creon-Datareader/blob/master/creonAPI.py

import sys
import time
import win32com.client


g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')


def check_PLUS_status(original_func):
    """original_func 콜하기 전에 PLUS 연결 상태 체크하는 데코레이터"""
    
    def wrapper(*args, **kwargs):
        
        bConnect = g_objCpStatus.IsConnect
        
        if bConnect == 0:
            print("PLUS is not properly connected...")
            sys.exit()
        
        return original_func(*args, **kwargs)
    
    return wrapper

# 서버로부터 과거의 차트 데이터 가져오는 클래스
class CpStockChart:
    
    def __init__(self):
        
        self.objStockChart = win32com.client.Dispatch('CpSysDib.StockChart')
        
    def _check_rq_status(self):
        """self.objStockChart.BlockRequest() 로 요청한 후 이 메소드로 통신상태 검사
        
        :return: None
        """
        
        rqStatus = self.objStockChart.GetDibStatus()
        rqRet = self.objStockChart.GetDibMsg1()
        
        if rqStatus == 0:
            # print("통신상태 정상[{}]{}".format(rqStatus, rqRet))
            pass
        else:
            print("통신상태 오류[{}]{} 종료합니다...".format(rqStatus, rqRet))
            sys.exit()
            
    @check_PLUS_status
    def request_dwm(self, code, dwm, count, caller: 'MainWindow', from_date=0, ohlcv_only=True):
        """차트 요청 - 최근 일부터 개수 기준
        
        :param code: 종목코드
        :param dwm: 'D': 일봉, 'W': 주봉, 'M': 월봉
        :param count: 요청할 데이터 개수
        :param caller: 이 메소드 호출한 인스턴스. 결과 데이터를 caller의 멤버로 전달
        
        :return: None
        """
        
        self.objStockChart.SetInputValue(0, code) # 종목코드
        self.objStockChart.SetInputValue(1, ord('2')) # 개수로 받기???
        self.objStockChart.SetInputValue(4, count) # 최근 count개
        
        if ohlcv_only:
            self.objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8]) # 요청항목 - 날짜, 시가, 고가, 저가, 종가, 거래량
            rq_column = ('date', 'open', 'high', 'low', 'close', 'volume')
        else:
            # 요청항목
            # - 날짜, 시가, 고가, 저가, 종가, 거래량, 
            # - 상장주식수, 외국인주문한도수량, 외국인현보유수량, 외국인현보유비율, 기관순매수, 기관누적순매수
            self.objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8, 12, 14, 16, 17, 20, 21])
            rq_columns = (
                'date', 'open', 'high', 'low', 'close', 'volume', 
                'num_listed_shares', 'foreigner_order_limit', 
                'foreigner_current_holdings', 'foreigner_current_holding_ratio', 
                'institutional_net_buying', 'institutional_cumulative_net_buying', 
            )
            
        self.objStockChart.SetInputValue(6, dwm) # 차트 주기 - 일/주/월
        self.objStockChart.SetInputValue(9, ord('1')) # 수정주가 사용
        
        rcv_data = {}
        for col in rq_columns:
            rcv_data[col] = []
            
        rcv_count = 0
        while count > rcv_count:
            self.objStockChart.BlockRequest() # 요청 후 응답 대기
            self._check_rq_status() # 통신상태 검사
            time.sleep(0.25) # 시간당 RQ 제한으로 인해 장애가 발생하지 않도록 딜레이 할당
            
            rcv_batch_len = self.objStockChart.GetHeaderValue(3) # 받아 온 데이터 개수
            rcv_batch_len = min(rcv_batch_len, count - rcv_count) # 정확히 count 개수만큼 받기 위함
            
            for i in range(rcv_batch_len):
                for col_idx, col in enumerate(rq_columns):
                    rcv_data[col].append(self.objStockChart.GetDataValue(col_idx, i))
                    
            if len(rcv_data['date']) == 0: # 데이터가 없는 경우
                print(code, "no data")
                return False
            
            # rcv_batch_len 만큼 받은 데이터의 가장 오래된 date
            rcv_oldest_date = rcv_data['date'][-1]
            
            rcv_count += rcv_batch_len
            caller.return_status_msg = '{} / {}'.format(rcv_count, count)
            
            # 서버가 가진 모든 데이터를 요청한 경우 break.
            # self.objStockChart.Continue 는 개수로 요청한 경우
            # count 만큼 이미 다 받았더라도 계속 1의 값을 가지고 있어서
            # while 조건문에서 count > rcv_count 를 체크
            if not self.objStockChart.Continue:
                break
            if rcv_oldest_date < from_date:
                break
            
        caller.rcv_data = rcv_data # 받은 데이터를 caller 의 멤버로 저장
        
        return True
    
    @check_PLUS_status
    def request_mt(self, code, dwm, tick_range, count, caller: 'MainWindow', from_date=0, ohlcv_only=True):
        """차트 요청 - 분간, 틱 차트
        
        :param code: 종목코드
        :param dwm: 'm': 분봉, 'T': 틱봉
        :param tick_range: 1분봉 or 5분봉, ...
        :param count: 요청할 데이터 개수
        :param caller: 이 메소드 호출한 인스턴스, 결과 데이터를 caller 의 멤버로 전달
        :return:
        """
        
        self.objStockChart.SetInputValue(0, code) # 종목코드
        self.objStockChart.SetInputValue(1, ord('2')) # 개수로 받기
        self.objStockChart.SetInputValue(4, count) # 조회 개수
        
        if ohlcv_only:
            self.objStockChart.SetInputValue(
                5, [0, 1, 2, 3, 4, 5, 8]) # 요청항목 - 날짜, 시가, 고가, 저가, 종가, 거래량
            rq_column = ('date', 'open', 'high', 'low', 'close', 'volume')
        else:
            # 요청항목
            # - 날짜, 시가, 고가, 저가, 종가, 거래량, 
            # - 상장주식수, 외국인주문한도수량, 외국인현보유수량, 외국인현보유비율, 기관순매수, 기관누적순매수
            self.objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8, 12, 14, 16, 17, 20, 21])
            rq_columns = (
                'date', 'open', 'high', 'low', 'close', 'volume', 
                'num_listed_shares', 'foreigner_order_limit', 
                'foreigner_current_holdings', 'foreigner_current_holding_ratio', 
                'institutional_net_buying', 'institutional_cumulative_net_buying', 
            )
            
        self.objStockChart.SetInputValue(6, dwm) # 차트 주기 - 분/틱
        self.objStockChart.SetInputValue(7, tick_range) # 분틱차트 주기
        self.objStockChart.SetInputValue(9, ord('1')) # 수정주가 사용
        
        rcv_data = {}
        for col in rq_columns:
            rcv_data[col] = []
            
        rcv_count = 0
        while count > rcv_count:
            self.objStockChart.BlockRequest() # 요청 후 응답 대기
            self._check_rq_status() # 통신상태 검사
            time.sleep(0.25) # 시간당 RQ 제한으로 인해 장애가 발생하지 않도록 딜레이 할당
            
            rcv_batch_len = self.objStockChart.GetHeaderValue(3) # 받아 온 데이터 개수
            rcv_batch_len = min(rcv_batch_len, count - rcv_count) # 정확히 count 개수만큼 받기 위함
            
            for i in range(rcv_batch_len):
                for col_idx, col in enumerate(rq_columns):
                    rcv_data[col].append(self.objStockChart.GetDataValue(col_idx, i))
                    
            if len(rcv_data['date']) == 0: # 데이터가 없는 경우
                print(code, "no data")
                return False
            
            # rcv_batch_len 만큼 받은 데이터의 가장 오래된 date
            rcv_oldest_date = int("{}{:04}".format(rcv_data['date'][-1], rcv_data['time'][-1]))
            
            rcv_count += rcv_batch_len
            caller.return_status_msg = "{} / {}(maximum)".format(rcv_count, count)
            
            # 서버가 가진 모든 데이터를 요청한 경우 break.
            # self.objStockChart.Continue 는 개수로 요청한 경우
            # count 만큼 이미 다 받았더라도 계속 1의 값을 가지고 있어서
            # while 조건문에서 count > rcv_count 를 체크
            if not self.objStockChart.Continue:
                break
            if rcv_oldest_date < from_date:
                break
            
        # 분몽의 경우 날짜와 시간을 하나의 문자열로 합친 후 int 로 변환
        rcv_data['date'] = list(map(
            lambda x, y: int("{}{:04}".format(x, y)), 
            rcv_data['date'], rcv_data['time']
        ))
        del rcv_data['time']
        caller.rcv_data = rcv_data # 받은 데이터를 caller 의 멤버로 저장
        
        return True
        
        
class CpCodeMgr:
    """종목코드 관리하는 클래스"""
    
    def __init__(self):
        
        self.objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
        
    def get_code_list(self, market):
        """마켓에 해당하는 종목코드 리스트 반환하는 메소드
        
        :param market: 1: kospi, 2:kosdaq, ...
        :return: market 에 해당하는 코드 list
        """
        
        return self.objCodeMgr.GetStockListByMarket(code)
    
    def get_section_code(self, code):
        """부구분코드를 반환하는 메소드"""
        
        return self.objCodeMgr.GetStockSectionKind(code)
    
    def get_code_name(self, code):
        """종목코드를 받아 종목명을 반환하는 메소드"""
        
        return self.objCodeMgr.CodeToName(code)