# -*- coding:utf-8 -*-
# author: LuGu
# initial date: 24/12/2021
# last edited date: 24/12/2021
# description: creon_datareader_cli
# reference: https://github.com/gyusu/Creon-Datareader/blob/master/creon_datareader_cli.py

import sys
import os
import gc
import pandas as pd
import tqdm
import sqlite3
import argparse

import creonAPI
from utils import is_market_open, available_latest_date, preformat_cjk


class CreonDatareaderCLI:
    
    def __init__(self):
        
        self.objStockChart = creonAPI.CpStockChart()
        self.objCodeMgr = creonAPI.CpCodeMgr()
        self.rcv_data = dict() # RQ 후 받아 온 데이터 자장 멤버
        
        self.sv_code_df = pd.DataFrame()
        self.db_code_df = pd.DataFrame()
        
        sv_code_list = self.objCodeMgr.get_code_list(1) + self.objCodeMgr.get_code_list(2)
        sv_name_list = list(map(self.objCodeMgr.get_code_name, sv_code_list))
        self.sv_code_df = pd.DataFrame({'stockCode': sv_code_list, 'stockName': sv_name_list})
        
    def update_price_db(self, db_fn, tick_unit='day', ohlcv_only=False):
        """
        :param db_fn: db 파일 경로.
        :param tick_unit: '1min', '5min', 'day'. 이미 db_fn 이 존재할 경우, 입력 값 무시하고 기존에 사용된 값 사용.
        :param ohlcv_only: ohlcv 이외의 데이터도 저장할지 여부. 이미 db_fn 이 존재할 경우, 입력 값 무시하고 기존에 사용된 값 사용
            'day' 아닌 경우 False 선택 불가 고정.
        """
        
        if tick_unit != 'day':
            ohlcv_only = True
            
        # 로컬 DB에 저장된 종목 정보 가져와서 dataframe으로 저장
        con = sqlite3.connect(db_fn)
        cursor = con.cursor()
        
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        db_code_list = cursor.fetchall()
        
        for i in range(len(db_code_list)):
            db_code_list[i] = db_code_list[i][0]
        
        db_name_list = list(map(self.objCodeMgr.get_code_name, db_code_list))
        
        db_latest_list = []
        for db_code in db_code_list:
            cursor.execute("SELECT date FROM {} ORDER BY date DESC LIMIT 1;".format(db_code))
            db_latest_list.append(cursor.fetchall()[0][0])
            
        # 현재 DB 에 저장된 'date' column 의 tick_unit 확인
        # 현재 DB 에 젖아된 column 명 확인. (ohlcv_only 여부 확인)
        if db_latest_list:
            cursor.execute("SELECT date FROM {} ORDER BY date ASC LIMIT 2;".format(db_code_list[0]))
            date0, date1 = cursor.fetchall()
            
            # 날짜가 분 단위인 경우
            if data0[0] > 99999999:
                if date1[0] - data0[0] == 5: # 5분 간격인 경우
                    tick_unit = '5min'
                else: # 1분 간격인 경우
                    tick_unit = '1min'
            elif data0[0] % 100 == 0: # 월봉인 경우
                tick_unit = 'month'
            elif data0[0] % 10 == 0: # 주봉인 경우
                tick_unit = 'week'
            else: # 일봉인 경우
                tick_unit = 'day'
                
            # column 개수로 ohlcv_only 여부 확인
            cursor.execute("SELECT * FROM {} LIMIT 1;".format(db_code_list[0]))
            column_names = [description[0] for description in cursor.description]
            if len(column_names) > 6: # date, o, h, l, c, v
                ohlcv_only = False
            else:
                ohlcv_only = True
                
        db_code_df = pd.DataFrame({
            'stockCode': db_code_list, 'stockName': db_name_list, 
            'updateDate': db_latest_list
        })
        fetch_code_df = self.sv_code_df
        
        # 분봉/일봉에 대해서만 아래 코드가 효과 존재
        if not is_market_open():
            latest_date = available_latest_date()
            
            if tick_unit == 'day':
                latest_date = latest_date // 10000
            
            # 이미 DB 데이터가 최신인 종목들은 가져올 목록에서 제외
            already_up_to_date_codes = db_code_df.loc[
                db_code_df['updateDate']==latest_date
            ]['stockCode'].values
            fetch_code_df = fetch_code_df.loc[
                fetch_code_df['stockCode'].apply(lambda x: x not in already_up_to_date_codes)
            ]
            
        if tick_unit == '1min':
            count = 200000 # 서버 데이터 최대 reach 약 18.5만 이므로 (18/02/25 기준)
            tick_range = 1
        elif tick_unit == '5min':
            count = 100000
            tick_range = 5
        elif tick_unit == 'day':
            count = 10000 # 10000 개면 현재부터 1980년 까지ㅣ의 데이터에 해당. 충분.
            tick_range = 1
        elif tick_unit == 'week':
            count = 2000
        elif tick_unit == 'month':
            count = 500
            
        if ohlcv_only:
            columns = ['open', 'high', 'low', 'close', 'volume']
        else:
            columns = [
                'open', 'high', 'low', 'close', 'volume', 
                'num_listed_shares', 'foreigner_order_limit', 
                'foreigner_current_holdings', 'foreigner_current_holding_ratio', 
                'institutional_net_buying', 'institutional_cumulative_net_buying', 
            ]
            
        with sqlite3.connect(db_fn) as con:
            cursor = con.cursor()
            tqdm_range = tqdm.trange(len(fetch_code_df), ncols=100)
            
            for i in tqdm_range:
                code = fetch_code_df.iloc[i]
                update_status_msg = "[{}] {}".format(code[0], code[1])
                tqdm_range.set_description(preformat_cjk(update_status_msg, 25))
                
                from_date = 0
                if code[0] in db_code_df['stockCode'].tolist():
                    cursor.execute("SELECT date FROM {} ORDER BY date DESC LIMIT 1;".format(code[0]))
                    from_date = cursor.fetchall()[0][0]
                    
                if tick_unit == 'day': # 일봉 데이터 받기
                    if self.objStockChart.request_dwm(
                        code[0], ord('D'), count, self, from_date, ohlcv_only
                    ) == False:
                        continue
                elif tick_unit == '1min' or tick_unit == '5min': # 분봉 데이터 받기
                    if self.objStockChart.request_mt(
                        code[0], ord('m'), tick_range, count, self, from_date, ohlcv_only
                    ) == False:
                        continue
                elif tick_unit == 'week': # 주봉 데이터 받기
                    if self.objStockChart.request_dwm(
                        code[0], ord('W'), count, self, from_date, ohlcv_only
                    ) == False:
                        continue
                elif tick_unit == 'month': # 월봉 데이터 받기
                    if self.objStockChart.request_dwm(
                        code[0], ord('M'), count, self, from_date, ohlcv_only
                    ) == False:
                        continue
                
                df = pd.DataFrame(self.rcv_data, columns=columns, index=self.rcv_data['date'])
                
                # 기존 DB 와 겹치는 부분 제거?
                if from_date != 0:
                    df = df.loc[:from_date]
                    df = df.iloc[:-1]
                    
                # 뒤집어서 저장 (결과적으로 date 기준 오름차순으로 저장)
                df = df.iloc[::-1]
                df.to_sql(code[0], con, if_exists='append', index_label='date')
                
                # 메모리 overflow 방지
                def df
                gc.collect()
                

def main_cli():
    
    p = argparse.ArgumentParser(description='creon datareader CLI')
    
    p.add_argument('--db_fn', required=True, type=str)
    p.add_argument('--tick_unit', required=False, type=str, default='day', help='{1min, 5min, day, week, month}')
    p.add_argument('--ohlcv_only', required=False, type=int, default=0, help='0: False, 1: True')
    args = p.parse_args()
    
    creon = CreonDatareaderCLI()
    creon.update_price_db(args.db_fn, args.tick_unit, args.ohlcv_only==1)
    
    
if __name__ == '__main__':
    main_cli()
    