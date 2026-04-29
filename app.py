import streamlit as st
import pandas as pd
import openpyxl
import io, tempfile, re
from openpyxl.styles import PatternFill # 엑셀 헤더 노란색으로 칠하기

st.set_page_config(page_title="더존 전표 변환기", page_icon="📊", layout="wide")
st.title("📊 더존 위하고 전표 변환기")


# ─────────────────────────────
# 숫자 변환
# ─────────────────────────────
def to_int(v):
    try:
        if v is None:
            return 0
        if isinstance(v, str):
            v = re.sub(r"[^0-9.-]", "", v)
        return int(float(v))
    except:
        return 0


def clean(v):
    if v is None:
        return ''
    try:
        if pd.isna(v):
            return ''
    except:
        pass
    return str(v).strip()


def parse_date(v):
    try:
        d = pd.to_datetime(v)
        return d.month, d.day
    except:
        return None, None
        
# ─────────────────────────────
# 🔥 거래유형 정규화 (핵심 추가)
# ─────────────────────────────
def normalize_trade_type(ttype):
    ttype = str(ttype)

    if "매도" in ttype:
        return "SELL"
    elif "매수" in ttype:
        return "BUY"
    elif "예탁금" in ttype or "예탁금이용료" in ttype:
        return "INTEREST"
    elif "입고" in ttype or "공모주입고" in ttype:
        return "StockCredit"
    elif "입금" in ttype or "이체입금" in ttype:
        return "Credit"
    elif "출금" in ttype and (
                              "은행이체" in ttype or
                              "타사이체" in ttype or
                              "이체출금" in ttype
                             ):
        return "Debit"
        
    return None

# ─────────────────────────────
# row (엑셀 10컬럼 기준)
# ─────────────────────────────
def row(m, d, div, acct_code, acct_name, cp_code, cp_name, memo, dr, cr):
     return [m, d, div, acct_code, acct_name, cp_code, cp_name, memo, dr, cr]
    
# ─────────────────────────────
# 🔥 수정 1: 종목명 추출 함수 추가
# ─────────────────────────────
def extract_stock_name(name):
    name = str(name).replace(" ", "").strip()
    if "#" in name:
        return name.split("#")[-1]
    return name


# ─────────────────────────────
# 🔥 수정 2: 매핑 구조 변경 (핵심)
# ─────────────────────────────
def load_broker_map(file):
    if file is None:
        return {}

    df = pd.read_excel(file)

    return {
        extract_stock_name(name): (   # 🔥 key = 종목명만
            str(code).strip(),       # 거래처코드
            str(name).strip()        # 원본 거래처명 유지
        )
        for code, name in zip(df.iloc[:, 0], df.iloc[:, 1])
    }


# ─────────────────────────────
# 🔥 수정 3: 매핑 조회 방식 변경
# ─────────────────────────────
def get_broker_info(stock, broker_map):

    key = extract_stock_name(stock)  # 🔥 동일한 기준으로 변환

    if key in broker_map:
        return broker_map[key]   # (code, name)
    else:
        return "", stock         # 미매핑


    
# ─────────────────────────────
# HANTOO 파서 (자동 컬럼 매칭)
# ─────────────────────────────
def parse_hantoo_sheet(df):
    header_row = None

    for i in range(min(15, len(df))):
        row_str = df.iloc[i].astype(str)
        if any("거래일" in str(v or "") for v in row_str): # v = 엑셀 한 셀 값 (문자/숫자/NaN 다 들어옴)
            header_row = i
            break

    if header_row is None:
        return []

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)

    def find_col(keys):
        for c in df.columns:
            for k in keys:
                if k in str(c):
                    return c
        return None
    
    c_date  = find_col(["거래일","거래일자","일자","날짜"])
    c_type  = find_col(["구분","적요명","내용","거래종류","거래명","거래구분","거래종류"])
    c_stock = find_col(["종목","종목명","종목명(거래상대명)","종목명(상대처)"])
    c_qty   = find_col(["수량","거래수량","거래좌수"])
    c_price = find_col(["단가","가격","거래단가","기준가"])
    c_net   = find_col(["금액","거래금액","입출금액","입금/입고/매도","출금/출고/매수","거래대금","입출금액"])
    c_fee   = find_col(["수수료/Fee","수수료"])
    c_tax   = find_col(["tax","세금","제세금","거래세/농특세","거래세"])

    trades = []

    for _, r in df.iterrows():
        try:
            m, d = parse_date(r.get(c_date))
            if not m:
                continue

            trade_type = clean(r.get(c_type))
            stock = clean(r.get(c_stock)).strip()

            qty = to_int(r.get(c_qty))
            price = to_int(r.get(c_price))
            net = to_int(r.get(c_net))
            fee = to_int(r.get(c_fee))
            tax = to_int(r.get(c_tax))

            if not trade_type:
                continue

            trades.append({
                "month": m,
                "day": d,
                "type": trade_type,
                "stock": stock,
                "qty": qty,
                "price": price,
                "net": net,
                "fee": fee,
                "tax": tax
            })

        except:
            continue

    return trades
    
# ─────────────────────────────
# 🔥 거래처 자동 매핑 함수
# ─────────────────────────────
def get_broker_code(stock, broker_map, default_code):
    return broker_map.get(stock, default_code)

# ─────────────────────────────
# 전표 생성
# ─────────────────────────────
#                                      거래처코드,   예치금,  단기매매증권,  이자수익,        배당금수익
def process_trades(trades, broker_map, broker_code, deposit, short_inv, interest_income, dividend_income):
    rows = []

    for t in trades:
        m = t["month"]
        d = t["day"]
        type = t["type"]      # 구분, 적요명,내용,거래종류
        stock = t["stock"]    # 종목명
        stock_name = extract_stock_name(stock)    #매핑용 거래처명
        ttype = normalize_trade_type(t["type"])
        qty = t["qty"]        # 수량
        price = t["price"]    # 단가
        net = t["net"]        # 거래금액
        fee = t["fee"]        # 거래수수료
        tax = t["tax"]        # 세금과공과

        # 🔥 여기!!!! (무조건 이 위치)
        ttype = normalize_trade_type(t["type"])

        if not ttype:
            continue
    
        
        # 🔥 매핑 적용
        cp_code, cp_name = get_broker_info(stock_name, broker_map)

        # 🔥 디버깅 (필요하면 주석 해제)
        # st.write("매핑확인:", stock, "→", cp_code, cp_name)

        # 매도
        if ttype == "SELL":
            memo = f"{stock_name}({qty}주*{price})매도"

            rows.append(row(m,d,"차변",deposit,"예치금",broker_code,"",memo,net,0))
            rows.append(row(m,d,"대변",short_inv,"단기매매증권",cp_code,cp_name,memo,0,qty*price))

        # 매수
        elif ttype == "BUY":
            cost = qty * price
            memo = f"{stock_name}({qty}주*{price})매수"

            rows.append(row(m,d,"차변",short_inv,"단기매매증권",cp_code,cp_name,memo,cost,0))
            rows.append(row(m,d,"차변",82800,"증권수수료",cp_code,cp_name,"매수수수료",fee,0))
            rows.append(row(m,d,"대변",deposit,"예치금",broker_code,"",memo,0,cost-fee))

        # 예탁금이용료
        elif ttype == "INTEREST":
            memo = "예탁금이용료"
        
            rows.append(row(m,d,"차변",deposit,"예치금",broker_code,"",memo,net,0))
            rows.append(row(m,d,"대변",interest_income,"이자수익(금융)",broker_code,stock,memo,0,net))

        # 공모주입고
        elif ttype == "StockCredit":
            cost = qty * price
            memo = f"{stock_name}({qty}주*{price})입고"

            rows.append(row(m,d,"차변",short_inv,"단기매매증권",cp_code,cp_name,memo,cost,0))
            rows.append(row(m,d,"대변",13100,"선급금",cp_code,cp_name,memo,0,cost))
    
        # 이체입금
        elif ttype == "Credit":
            memo = f"{type}"

            rows.append(row(m,d,"차변",deposit,"예치금",broker_code,"",memo,net,0))
            rows.append(row(m,d,"대변",deposit,"예치금","","미등록거래처",memo,0,net))
    
        # 이체출금
        elif ttype == "Debit":
            memo = f"{type}"

            rows.append(row(m,d,"차변",deposit,"예치금","","미등록거래처",memo,0,net))
            rows.append(row(m,d,"대변",deposit,"예치금",broker_code,"",memo,net,0))
    return rows

# ─────────────────────────────
# Excel 생성
# ─────────────────────────────
def create_excel(rows):
    wb = openpyxl.Workbook()
    ws = wb.active

    header = [
        "1.월","2.일","3.구분",
        "4.계정과목코드","5.계정과목명",
        "6.거래처코드","7.거래처명",
        "8.적요명","9.차변(출금)","10.대변(입금)"
    ]

    ws.append(header)

    # 🔥 노란색 스타일
    yellow_fill = PatternFill(
        start_color="FFF59D",  # 연노랑
        end_color="FFF59D",
        fill_type="solid"
    )

    # 헤더 스타일 적용
    for col in range(1, len(header) + 1):
        ws.cell(row=1, column=col).fill = yellow_fill

    # 데이터 입력
    for r in rows:
        ws.append(r)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# ─────────────────────────────
# UI
# ─────────────────────────────
# 🔥 예치금 (사용자 입력)
st.text_input("예치금", placeholder="예: 12500")
# 🔥 단기매매증권 (사용자 입력)
short_inv = st.text_input("단기매매증권", placeholder="예: 10700")
# 🔥 이자수익(금융) (사용자 입력)
interest_income = st.text_input("이자수익(금융)", placeholder="예: 42000"))
# 🔥 배당금수익 (사용자 입력)
dividend_income = st.text_input("배당금수익", placeholder="예: 41800"))

# 🔥 거래처 매핑 엑셀
broker_file = st.file_uploader("거래처 매핑 엑셀 (이름 / 코드)", type=["xlsx"])

# 🔥 거래처코드(금융사) (사용자 입력)
broker_code = st.text_input("증권사 거래처코드")

# 🔥 엑셀 파일 업로드
uploaded = st.file_uploader("엑셀 업로드")

if uploaded:

    if st.button("변환 실행"):

        broker_map = load_broker_map(broker_file) if broker_file else {}

        file_bytes = uploaded.getvalue()

        try:
            xl = pd.ExcelFile(io.BytesIO(file_bytes))
        except Exception as e:
            st.error(f"엑셀 읽기 실패: {e}")
            st.stop()

        all_trades = []

        for sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet, header=None)
            trades = parse_hantoo_sheet(df)
            all_trades.extend(trades)

        st.write("총 trades:", len(all_trades))
#                                                      거래처코드,   예치금,  단기매매증권,  이자수익,        배당금수익
        rows = process_trades(all_trades, broker_map, broker_code, deposit, short_inv, interest_income, dividend_income)

        if not rows:
            st.error("❌ 변환 데이터 없음")
        else:
            out = create_excel(rows)

            st.success("완료")
            st.download_button(
                "다운로드",
                data=out,
                file_name="result.xlsx"
            )
