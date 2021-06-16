# link : https://airfox1.tistory.com/2?category=1118519

import time, win32con, win32api, win32gui, ctypes #pip3 install pypiwin32
from pywinauto import clipboard #pip3 install pywinauto , pandas
import pandas as pd #채팅내용을 DataFrame 형태로 확인

# 세팅값 지정
kakao_opentalk_name = '헌'
window_classname = "RICHEDIT50W"
chat_command = '명령'
chat_command2 = '배고파'

# 전역변수 지정
w = win32con





# 조합키 사용
PBYTE256 = ctypes.c_ubyte * 256
_user32 = ctypes.WinDLL("user32")
GetKeyboardState = _user32.GetKeyboardState
SetKeyboardState = _user32.SetKeyboardState

PostMessage = win32api.PostMessage
SendMessage = win32gui.SendMessage

GetCurrentThreadId = win32api.GetCurrentThreadId
GetWindowThreadProcessId = _user32.GetWindowThreadProcessId
AttachThreadInput = _user32.AttachThreadInput
MapVirtualKeyA = _user32.MapVirtualKeyA

MakeLong = win32api.MAKELONG
IsWindow = win32gui.IsWindow

def PostKeyEx(hwnd, key, shift, specialkey):

    if IsWindow(hwnd):
        ThreadId = GetWindowThreadProcessId(hwnd, None)

        lparam = MakeLong(0, MapVirtualKeyA(key, 0))
        msg_down = w.WM_KEYDOWN
        msg_up = w.WM_KEYUP

        if specialkey:
            lparam = lparam | 0x1000000

        if len(shift) > 0:  # Если есть модификаторы - используем PostMessage и AttachThreadInput
            pKeyBuffers = PBYTE256()
            pKeyBuffers_old = PBYTE256()

            SendMessage(hwnd, w.WM_ACTIVATE, w.WA_ACTIVE, 0)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, True)
            GetKeyboardState(ctypes.byref(pKeyBuffers_old))

            for modkey in shift:
                if modkey == w.VK_MENU:
                    lparam = lparam | 0x20000000
                    msg_down = w.WM_SYSKEYDOWN
                    msg_up = w.WM_SYSKEYUP
                pKeyBuffers[modkey] |= 128

            SetKeyboardState(ctypes.byref(pKeyBuffers))
            time.sleep(0.01)
            PostMessage(hwnd, msg_down, key, lparam)
            time.sleep(0.01)
            PostMessage(hwnd, msg_up, key, lparam | 0xC0000000)
            time.sleep(0.01)
            SetKeyboardState(ctypes.byref(pKeyBuffers_old))
            time.sleep(0.01)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, False)

        else:  # Если нету модификаторов - используем SendMessage
            SendMessage(hwnd, msg_down, key, lparam)
            SendMessage(hwnd, msg_up, key, lparam | 0xC0000000)





#지정 문구 전송
def kakao_sendtext(text):
    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    SendReturn(hwndEdit)

#엔터
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

#채팅내용 복사
def kakao_copy():
    PostKeyEx(hwndListControl, ord('A'), [w.VK_CONTROL], False)
    time.sleep(1)
    PostKeyEx(hwndListControl, ord('C'), [w.VK_CONTROL], False)
    ttext = clipboard.GetData()
    return ttext


#채팅방 내용 저장 - 초기값 지정
def chat_last_save():
    ttext = kakao_copy()
    a = ttext.split('\r\n')   # \r\n 으로 스플릿 __ 대화내용 인용의 경우 \r 때문에 해당안됨
    df = pd.DataFrame(a)    # DF 으로 바꾸기

    df[0] = df[0].str.replace('\[([\S\s]+)\] \[(오전|오후)([0-9:\s]+)\] ', '')  # 정규식으로 채팅내용만 남기기
    return df.index[-2], df.iloc[-2, 0]

#채팅방 내용 저장 - 주기적 확인
def chat_chek_command(cls, clst):
    ttext = kakao_copy()
    a = ttext.split('\r\n')  # \r\n 으로 스플릿 __ 대화내용 인용의 경우 \r 때문에 해당안됨
    df = pd.DataFrame(a)  # DF 으로 바꾸기

    df[0] = df[0].str.replace('\[([\S\s]+)\] \[(오전|오후)([0-9:\s]+)\] ', '')  # 정규식으로 채팅내용만 남기기

    if df.iloc[-2, 0] == clst:
        print("신규채팅 없음")
        return df.index[-2], df.iloc[-2, 0]
    else:
        print("신규채팅 발생")
        df1 = df.iloc[cls+1 : , 0]   # 최근 채팅내용만 남김

        found = df1[ df1.str.contains(chat_command) ]    # 명령어 여부 카운트
        found2 = df1[ df1.str.contains(chat_command2) ]    # 명령어 여부 카운트

        if 1 <= int(found.count()):
            print("<--[커맨드 확인]-->")
            kakao_sendtext("삐빅! 감지되었습니다.")  # 메시지 전송,, 추후 함수로 분리하여 처리

            # 명령어 여러개 쓸경우 리턴값으로 각각 빼서 쓰면 될듯. 일단 테스트용으로 위에 하드코딩 해둠
            return df.index[-2], df.iloc[-2, 0]

        if 1 <= int(found2.count()):
            print("<--[커맨드 확인]-->")
            kakao_sendtext("삐빅! 자동응답기. 밥머그에숑")  # 메시지 전송,, 추후 함수로 분리하여 처리

            # 명령어 여러개 쓸경우 리턴값으로 각각 빼서 쓰면 될듯. 일단 테스트용으로 위에 하드코딩 해둠
            return df.index[-2], df.iloc[-2, 0]

        else:
            print("<--[커맨드 없음]-->")
            return df.index[-2], df.iloc[-2, 0]

#채팅방 열기
def open_chatroom(chatroom_name):
    # # 채팅방 목록 검색하는 Edit (채팅방이 열려있지 않아도 전송 가능하기 위하여)
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    hwndkakao_edit1 = win32gui.FindWindowEx( hwndkakao, None, "EVA_ChildWindow", None)
    hwndkakao_edit2_1 = win32gui.FindWindowEx( hwndkakao_edit1, None, "EVA_Window", None)
    hwndkakao_edit2_2 = win32gui.FindWindowEx( hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)
    hwndkakao_edit3 = win32gui.FindWindowEx( hwndkakao_edit2_2, None, "Edit", None)

    # # Edit에 검색 _ 입력되어있는 텍스트가 있어도 덮어쓰기됨
    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(1)   # 안정성 위해 필요
    SendReturn(hwndkakao_edit3)
    time.sleep(1)











open_chatroom(kakao_opentalk_name)  # 채팅방 열기

# 핸들러
hwndMain = win32gui.FindWindow( None, kakao_opentalk_name)
hwndEdit = win32gui.FindWindowEx( hwndMain, None, window_classname, None)
hwndListControl = win32gui.FindWindowEx( hwndMain, None, "EVA_VH_ListControl_Dblclk", None)

#핸들러 정보 확인
print(hwndMain)
print(hwndEdit)
print(hwndListControl)

cls, clst = chat_last_save()  # 초기설정 _ 마지막채팅 저장
while True:
    print(".")
    cls, clst = chat_chek_command(cls, clst)  # 커멘드 체크
    time.sleep(5)
    
    win32gui.SendMessage(hwndMain, win32con.WM_CLOSE, 0, 0) #창 닫기
