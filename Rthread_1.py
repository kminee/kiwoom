from PyQt5.QtCore import *         # thread에서 사용할 함수를 불러오자
from kiwoom import Kiwoom          # 로그인을 하려면 Kiwoom함수가 필요하다
from PyQt5.QtWidgets import *      # 기능사용을 위해 불러온다

class Thread1(QThread):
    def __init__(self, parent):   # 메인 화면, 즉 부모의 윈도우 창을 가져오자
        super().__init__(parent)  # 부모의 윈도우 창을 초기화 하고 시작하자
        self.parent = parent      # 부모의 윈도우를 사용하기 위해 작업해야지


        #처음에는 키움증권 함수르 사용을 해야하니 Kiwoom 능력 상속 필요.
        self.k = Kiwoom()
        ##################

        ##사용되는 변수를 정리하자
        self.Acc_Screen = "1000"         # 계좌평가잔고내역을 받기위한 스크린

        ###### 슬롯
        self.k.kiwoom.OnReceiveTrData.connect(self.trdata_slot)  #  #trdata_slot을 설정하고 여기에 데이터를 던진다
        ###### EventLoop

        self.detail_account_info_event_loop = QEventLoop()  # 계좌 이벤트루프

        ###### 계좌정보 가져오기
        self.getItemList()               # 종목 이름 받아오기
        self.detail_acount_mystock()     # 계좌평가잔고내역 가져오기


    def getItemList(self):
        marketList = ["0", "10"]

        for market in marketList:
            codeList = self.k.kiwoom.dynamicCall("GetCodeListByMarket(QString)", market).split(";")[:-1]

            for code in codeList:
                name = self.k.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)
                self.k.All_Stock_Code.update({code: {"종목명": name}})

    def detail_acount_mystock(self, sPrevNext="0"):

        print("계좌평가잔고내역 조회")
        account = self.parent.redaccComboBox.currentText()  # ui에서 생성했던 콤보박스 안에서 가져오는 부분
        self.account_num = account
        print("최종 선택 계좌는 %s" % self.account_num)

        self.k.kiwoom.dynamicCall("SetInputValue(String, String)", "계좌번호", account)
        self.k.kiwoom.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")  # 모의투자 0000
        self.k.kiwoom.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.k.kiwoom.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.k.kiwoom.dynamicCall("CommRqData(String, String, int, String)", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.Acc_Screen)
        self.detail_account_info_event_loop.exec_()


    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):

        if sRQName == "계좌평가잔고내역요청":

            column_head = ["종목번호", "종목명", "보유수량", "매입가", "현재가", "평가손익", "수익률(%)"]
            colCount = len(column_head)
            rowCount = self.k.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            self.parent.redstockTableWidget_1.setColumnCount(colCount)                 # 행 갯수
            self.parent.redstockTableWidget_1.setRowCount(rowCount)                    # 열 갯수 (종목 수)
            self.parent.redstockTableWidget_1.setHorizontalHeaderLabels(column_head)   # 행의 이름 삽입

            print("계좌에 들어있는 종목 수 %s" % rowCount)

            totalBuyingPrice = int(self.k.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총매입금액"))
            currentTotalPrice = int(self.k.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총평가금액"))
            balanceAsset = int(self.k.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "추정예탁자산"))
            totalEstimateProfit = int(self.k.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총평가손익금액"))
            total_profit_loss_rate = float(self.k.kiwoom.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총수익률(%)"))

            #################################### 텍스트 라벨에 집어 넣기
            ##콤마 추가

            self.parent.totalBuyingPriceLabel.setText(str(format(totalBuyingPrice, ",")))
            self.parent.currentTotalPriceLabel.setText(str(format(currentTotalPrice, ",")))
            self.parent.balanceAssetLabel.setText(str(format(balanceAsset, ",")))
            self.parent.totalEstimateProfitLabel.setText(str(format(totalEstimateProfit, ",")))
            self.parent.totalpercenLabel.setText(str(format(total_profit_loss_rate, ",")))

            #################################################################


            for index in range(rowCount):
                itemCode = self.k.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, index, "종목번호").strip(" ").strip("A")
                itemName = self.k.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, index, "종목명")
                amount = int(self.k.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, index, "보유수량"))
                buyingPrice = int(self.k.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, index, "매입가"))
                currentPrice = int(self.k.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, index, "현재가"))
                estimateProfit = int(self.k.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, index, "평가손익"))
                profitRate = float(self.k.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, index, "수익률(%)"))
                total_chegual_price = self.k.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, index, "매입금액")
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = self.k.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, index, "매매가능수량")
                possible_quantity = int(possible_quantity.strip())

                if itemCode in self.k.acc_portfolio:
                    pass
                else:
                    self.k.acc_portfolio.update({itemCode:{}})      # self.account_stock_dict[code] = {}

                self.k.acc_portfolio[itemCode].update({"종목명": itemName.strip()})
                self.k.acc_portfolio[itemCode].update({"보유수량": amount})
                self.k.acc_portfolio[itemCode].update({"매입가": buyingPrice})
                self.k.acc_portfolio[itemCode].update({"수익률(%)": profitRate})
                self.k.acc_portfolio[itemCode].update({"현재가": currentPrice})
                self.k.acc_portfolio[itemCode].update({"매입금액": total_chegual_price})
                self.k.acc_portfolio[itemCode].update({"매매가능수량": possible_quantity})

                self.parent.redstockTableWidget_1.setItem(index, 0, QTableWidgetItem(str(itemCode)))
                self.parent.redstockTableWidget_1.setItem(index, 1, QTableWidgetItem(str(itemName)))
                ##천단위 구분 콤마
                self.parent.redstockTableWidget_1.setItem(index, 2, QTableWidgetItem(str(format(amount, ","))))
                self.parent.redstockTableWidget_1.setItem(index, 3, QTableWidgetItem(str(format(buyingPrice, ","))))
                self.parent.redstockTableWidget_1.setItem(index, 4, QTableWidgetItem(str(format(currentPrice, ","))))
                self.parent.redstockTableWidget_1.setItem(index, 5, QTableWidgetItem(str(format(estimateProfit, ","))))
                self.parent.redstockTableWidget_1.setItem(index, 6, QTableWidgetItem(str(format(profitRate, ","))))

            if sPrevNext == "2":
                self.detail_acount_mystock(sPrevNext="2")  # 다음 페이지가 있으면 전부 검색한다.
            else:
                self.detail_account_info_event_loop.exit()  # 끊어 준다.









