# -*- coding:utf-8 -*-
# author: LuGu
# initial date: 04/12/2021
# last edited date: 04/12/2021
# description: (테스트) Slack으로 주가를 조회하여 전송하는 봇 만들기

import os
import json
import requests
from datetime import datetime


NOW = datetime.now().strftime('%Y.%m.%d %H:%M:%S')
CUR_DIR = os.path.dirname(os.path.realpath(__file__))
PAR_DIR = os.path.dirname(CUR_DIR)
JSON_SlACK_PATH = os.path.join(CUR_DIR, 'src', 'slack_token.json')


def get_slack_token():
    with open(JSON_SlACK_PATH, 'r') as f:
        slack_token = json.load(f)['token']
    return slack_token

def post_message(channel='#tradingsystem', text=''):
    slack_token = get_slack_token()
    headers = {
        'Content-Type': 'application/json', 
        'Authorization': 'Bearer ' + slack_token
    }
    payload = {'channel': channel, 'text': text}
    res = requests.post(
        'https://slack.com/api/chat.postMessage', 
        headers=headers, data=json.dumps(payload)
    )
    return None
    
    
if __name__ == '__main__':
    post_message(text=f'{NOW} TEST')
    