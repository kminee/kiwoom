import sys
import json
import time
from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop, QTimer  # QEventLoop와 QTimer 임포트


## 일봉 데이터 수집기 클래스
class Kiwoom:
    def __init__(self):
        # QApplication 인스턴스는 한 번만 생성하고, 애플리케이션의 생명주기 동안 유지되어야 합니다.
        # sys.argv가 필요 없으므로 빈 리스트를 전달합니다.
        self.app = QApplication([])
        self.ocx = QAxWidget('KHOPENAPI.KHOpenAPICtrl.1')

        # 이벤트 핸들 연결
        self.ocx.OnEventConnect.connect(self._on_login)
        self.ocx.OnReceiveTrData.connect(self._on_receive_tr_data)
        # TR 요청 시 API 제한 등으로 인해 일정 시간 대기해야 할 경우를 대비해 타이머 설정
        self.event_loop = QEventLoop()
        self.tr_event_loop = QEventLoop()  # TR 데이터 수신을 위한 별도의 이벤트 루프

        self.login_ok = False
        self.tr_data = None
        self.data_ready = False
        self.current_rqname = ""  # 현재 요청 중인 TR의 rqname을 저장

    def connect(self):
        """키움증권 API에 로그인 시도 및 로그인 완료까지 대기합니다."""
        print(">> 로그인 시도...")
        self.ocx.dynamicCall("CommConnect()")
        # 로그인 완료 시그널을 받을 때까지 이벤트 루프를 블로킹하여 대기합니다.
        self.event_loop.exec_()
        if self.login_ok:
            print(">> 로그인 성공")
        else:
            print(">> 로그인 실패 또는 시간 초과")
            sys.exit(1)  # 로그인 실패 시 애플리케이션 종료

    def _on_login(self, err_code):
        """로그인 결과에 따라 처리합니다."""
        if err_code == 0:
            self.login_ok = True
        else:
            self.login_ok = False
            print(f"로그인 오류 발생: {err_code}")
        self.event_loop.exit()  # 로그인 이벤트가 발생하면 대기 중인 이벤트 루프를 종료합니다.

    def get_code_list(self):
        """코스피 및 코스닥 종목 코드를 가져옵니다."""
        kospi = self.ocx.dynamicCall("GetCodeListByMarket(QString)", ["0"]).split(";")
        kosdaq = self.ocx.dynamicCall("GetCodeListByMarket(QString)", ["10"]).split(";")
        codes = list(filter(None, kospi + kosdaq))
        print(f">> 총 종목 수: {len(codes)}")
        return codes

    def get_stock_name(self, code):
        """종목 코드로 종목명을 가져옵니다."""
        return self.ocx.dynamicCall("GetMasterCodeName(QString)", [code])

    def is_valid_stock(self, code, name):
        """유효한 종목인지 확인합니다."""
        keywords = ['우', 'ETF', 'ETN', '리츠', '스팩', '채권', '선물', '옵션', 'ELS']  # 제외할 키워드 추가
        if any(k in name for k in keywords):
            return False
        # 거래정지 확인
        status = self.ocx.dynamicCall("GetMasterStockState(QString)", [code])
        if "정지" in status or "거래정지" in status:  # "거래정지" 문자열도 확인
            return False
        return True

    def get_ohlcv(self, code):
        """특정 종목의 일봉 데이터를 요청하고 수신될 때까지 대기합니다."""
        self.tr_data = []
        self.data_ready = False
        self.current_rqname = "opt10081_req"  # 현재 요청하는 TR의 rqname 설정

        self.ocx.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "기준일자", "")  # 빈 문자열은 오늘 날짜를 의미합니다.
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "수정추가구분", "1")  # 1: 수정주가 반영

        # CommRqData 호출
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)", self.current_rqname, "opt10081", 0, "0101")

        # TR 데이터 수신 완료 시그널을 받을 때까지 대기
        # 이 부분이 QEventLoop를 사용하는 핵심입니다.
        self.tr_event_loop.exec_()
        return self.tr_data

    def _on_receive_tr_data(self, screen_no, rqname, trcode, recordname, prev_next):
        """TR 데이터 수신 시 처리합니다."""
        # 현재 요청한 TR인지 확인
        if rqname != self.current_rqname:
            return

        count = self.ocx.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        result = []

        for i in range(count):
            # GetCommData 호출 시 'recordname'은 무시하고, 'data_name'과 'index'만 사용합니다.
            # 데이터 타입을 int로 바로 변환하여 저장합니다.
            date = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, "", i, "일자").strip()
            open_ = int(
                self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, "", i, "시가").strip())
            high = int(self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, "", i, "고가").strip())
            low = int(self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, "", i, "저가").strip())
            close = int(
                self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, "", i, "현재가").strip())
            volume = int(
                self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, "", i, "거래량").strip())

            result.append({
                "date": date,
                "open": open_,
                "high": high,
                "low": low,
                "close": close,
                "volume": volume
            })

        self.tr_data = result[:240]  # 최근 240일치 데이터만 저장 (필요에 따라 조절)
        self.data_ready = True
        self.tr_event_loop.exit()  # TR 데이터 수신 완료 시 대기 중인 TR 이벤트 루프를 종료합니다.


### 채널을 작동하는 스크립트
if __name__ == "__main__":
    kiwoom = Kiwoom()
    kiwoom.connect()  # 로그인 및 API 연결

    all_data = {}
    codes = kiwoom.get_code_list()

    for idx, code in enumerate(codes):
        name = kiwoom.get_stock_name(code)
        if not kiwoom.is_valid_stock(code, name):
            print(f"[{idx + 1}/{len(codes)}] {code} {name} (제외: 유효하지 않은 종목)")
            continue

        print(f"[{idx + 1}/{len(codes)}] {code} {name} 데이터 요청 중...")
        try:
            ohlcv = kiwoom.get_ohlcv(code)
            all_data[code] = {
                "name": name,
                "ohlcv": ohlcv
            }
        except Exception as e:
            print(f"  >> {code} {name} 실패: {e}")

        # TR 요청은 초당 5회로 제한됩니다. 0.3초는 TR 당 약 3.3회로 안전합니다.
        time.sleep(0.3)

    with open("all_stock_data.json", "w", encoding="utf-8") as f:
        json.dump(all_data, f, indent=2, ensure_ascii=False)

    print("\n 모든 데이터 저장 완료: all_stock_data.json")

    # 모든 작업이 끝난 후 QApplication의 메인 이벤트 루프를 종료합니다.
    # 이 스크립트는 모든 데이터를 수집한 후 종료되므로, 이 부분은 일반적으로 필요하지 않을 수 있습니다.
    # 하지만 GUI 없이 백그라운드에서 실행되는 PyQt 애플리케이션의 경우 명시적으로 종료하는 것이 좋습니다.
    # sys.exit(kiwoom.app.exec_())
    # 모든 데이터 수집 후 프로그램이 자동으로 종료되도록 합니다.
    kiwoom.app.quit()