# link : https://airfox1.tistory.com/2?category=1118519
import time, win32con, win32api, win32gui  #pip3 install pypiwin32
import pandas as pd #가져온 채팅내용 DF로 쓸거라서
from module import PostKeyEx as pk
from pywinauto import clipboard #pip3 install pywinauto , pandas

# # 카톡창 이름 (열려있는 상태, 최소화 X, 창뒤에 숨어있는 비활성화 상태 가능)
kakao_opentalk_name = '채팅방이름'
chat_command = '명령'





# # 채팅내용 초기 저장 _ 마지막 채팅
def chat_last_save():
    pk.PostKeyEx(hwndListControl, ord('A'), [w.VK_CONTROL], False)
    time.sleep(1)
    pk.PostKeyEx(hwndListControl, ord('C'), [w.VK_CONTROL], False)
    ttext = clipboard.GetData()

    a = ttext.split('\r\n')   # \r\n 으로 스플릿 __ 대화내용 인용의 경우 \r 때문에 해당안됨
    df = pd.DataFrame(a)    # DF 으로 바꾸기

    df[0] = df[0].str.replace('\[([\S\s]+)\] \[(오전|오후)([0-9:\s]+)\] ', '')  # 정규식으로 채팅내용만 남기기

    return df.index[-2], df.iloc[-2, 0]


# # 채팅방 커멘드 체크
def chat_chek_command(cls, clst):
    pk.PostKeyEx(hwndListControl, ord('A'), [w.VK_CONTROL], False)
    time.sleep(1)
    pk.PostKeyEx(hwndListControl, ord('C'), [w.VK_CONTROL], False)
    ttext = clipboard.GetData()

    a = ttext.split('\r\n')  # \r\n 으로 스플릿 __ 대화내용 인용의 경우 \r 때문에 해당안됨
    df = pd.DataFrame(a)  # DF 으로 바꾸기

    df[0] = df[0].str.replace('\[([\S\s]+)\] \[(오전|오후)([0-9:\s]+)\] ', '')  # 정규식으로 채팅내용만 남기기

    if df.iloc[-2, 0] == clst:
        print("채팅 없었음..")
        return df.index[-2], df.iloc[-2, 0]
    else:
        print("채팅 있었음")

        df1 = df.iloc[cls+1 : , 0]   # 최근 채팅내용만 남김

        found = df1[ df1.str.contains(chat_command) ]    # 챗 카운트

        if 1 <= int(found.count()):
            print("-------커멘드 확인!")
            kakao_sendtext("답장!")  # 메시지 전송, time/실검

            # 명령어 여러개 쓸경우 리턴값으로 각각 빼서 쓰면 될듯. 일단 테스트용으로 위에 하드코딩 해둠
            return df.index[-2], df.iloc[-2, 0]

        else:
            print("커멘드 미확인")
            return df.index[-2], df.iloc[-2, 0]





# # 지정 문구를 전송
def kakao_sendtext(text):
    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    SendReturn(hwndEdit)

# # 엔터
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

# # 핸들
hwndMain = win32gui.FindWindow( None, kakao_opentalk_name)
hwndEdit = win32gui.FindWindowEx( hwndMain, None, "RICHEDIT50W", None)
hwndListControl = win32gui.FindWindowEx( hwndMain, None, "EVA_VH_ListControl_Dblclk", None)

# # 채팅 전송
text = "==시작=="
kakao_sendtext(text)

# # 정보 확인
print(hwndMain)
print(hwndEdit)
print(hwndListControl)

# #조합키, 본문을 클립보드에 복사 ( ctl + A , C )
PostKeyEx(hwndListControl, ord('A'), [w.VK_CONTROL], False)
time.sleep(1)
PostKeyEx(hwndListControl, ord('C'), [w.VK_CONTROL], False)
ctext = clipboard.GetData()
print(ctext)	# 내용 확인

cls, clst = chat_last_save()  # 초기설정 _ 마지막채팅 저장
while True:
    print("실행중.................")
    cls, clst = chat_chek_command(cls, clst)  # 커멘드 체크
    time.sleep(5)
