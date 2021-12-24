# -*- coding:utf-8 -*-
# author: LuGu
# initial date: 24/12/2021
# last edited date: 24/12/2021
# description: creon_datareader_cli
# reference: https://github.com/gyusu/Creon-Datareader/blob/master/utils.py

from datetime import datetime, timedelta
import unicodedata


def is_market_open():
    """
    :return: True: 장 중, False: 장 개장 전 or 마감 후
    """
    
    now = datetime.now()
    mmhh = int("{}{:02}".format(now.hour, now.minute))
    
    if mmhh < 900 or mmhh > 1530:
        return False
    if now.weekday() >= 5: # 토, 일
        return False
    
    return True


def available_latest_date():
    
    now = datetime.now()
    mmhh = int("{}{:02}".format(now.hour, now.minute))
    
    # 장 중에는 최신 데이터 연속적으로 발생하므로 None 반환
    if is_market_opne():
        return None
    
    # 주말인 경우 (이외의 공휴일 체크 구현 필요)
    if now.weekday() >= 5:
        latest_date = latest_date - timedelta(days=2)
        return cvt_dt_to_int(latest_date)
    
    # 주 중인 경우
    if mmhh > 1530: # 장 마감 후
        return cvt_dt_to_int(latest_date)
    else: # 장 개장 전
        latest_date = latest_date - timedelta(days=1)
        if latest_date.weekday() == 6:
            latest_date = latest_date - timedelta(days=2)
        return cvt_dt_to_int(latest_date)
    
    
def cvt_dt_to_int(dttm):
    
    return int(dttm.strftime('%Y%m%d%H%M'))


def preformat_cjk(string, width, align='<', fill=' '):
    
    count = (width - sum(1 + (unicodedata.east_asian_width(c) in 'MK') for c in string))
    
    return {
        '>': lambda s: fill * count + s, 
        '<': lambda s: s + fill * count, 
        '^': lambda s: fill * (count / 2) + s + fill * (count / 2 + count % 2)
    }[align](string)