# https://choi97201.tistory.com/m/32
import os
import win32com.client
from pywinauto import application
import time


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
                id='아이디', pwd='패스워드', pwdcert='공인인증서패스워드')
        )
        # 연결 될때까지 무한루프
        while True:
            if CpCybos.IsConnect:
                break
            time.sleep(1)

        print('connected.')
    return CpCybos


# 이미 연결되어있다면 재연결 x
CpCybos = connect(False)