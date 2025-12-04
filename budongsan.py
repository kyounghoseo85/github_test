# -*- coding: utf-8 -*-
import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook

# ===================================================================
# 1. 입력 데이터 (이 값들을 원하는 대로 변경하세요)
# ===================================================================
# --- 시간 정보 ---
purchase_date_str = "2021-07-01"  # 매매 시점 (1일로 통일)
current_date_str = "2025-12-02"  # 현재 시점

# --- 매매 정보 ---
purchase_price = 470_000_000
current_price = 370_000_000

# --- 중개수수료율 ---
def get_brokerage_rate(price):
    """부동산 거래 중개수수료 계산 (5억 미만: 0.4%, 5억 이상: 0.3%)"""
    if price < 450_000_000:
        return 0.4
    else:
        return 0.3

# --- 대출 정보 ---
loan_principal = 300_000_000
annual_interest_rate = 2.75  # 연 이자율 (%)
loan_term_years = 30

# --- 초기 투자 비용 ---
# 매매 중개수수료 계산 (동적)
brokerage_fee_rate = get_brokerage_rate(purchase_price)
brokerage_fee_buy = purchase_price * (brokerage_fee_rate / 100)

acquisition_tax = 5_170_000
interior_cost = 15_000_000

# --- 연간 보유 비용 ---
property_tax_per_year = 103_320  # 1년에 2번 납부 (상반기, 하반기)

# --- 미래 시나리오 (5년 후) ---
scenario_months = 60  # 분석 기간 (개월)
scenario_sale_price = 425_000_000 # 예상 매도가 (이 값은 엑셀에서 변경 가능)


# ===================================================================
# 2. 계산 로직 (원금균등 상환 방식)
# ===================================================================

# --- 날짜 및 기간 계산 ---
purchase_date = datetime.strptime(purchase_date_str, "%Y-%m-%d")
current_date = datetime.strptime(current_date_str, "%Y-%m-%d")
# 총 경과 개월 수 계산
elapsed_months = (current_date.year - purchase_date.year) * 12 + (current_date.month - purchase_date.month)

# --- Sheet 1: 데이터 입력 ---
initial_cash_investment = purchase_price - loan_principal
initial_investment_costs = initial_cash_investment + acquisition_tax + brokerage_fee_buy + interior_cost

data_input_data = [
    ["구분", "값", "비고"],
    ["[ 시간 정보 ]", "", ""],
    ["매매 시점", purchase_date_str, ""],
    ["현재 시점", current_date_str, f"총 보유 기간: {elapsed_months}개월"],
    ["[ 매매 정보 ]", "", ""],
    ["매매가", purchase_price, ""],
    ["현재 시세", current_price, ""],
    ["[ 대출 정보 ]", "", ""],
    ["대출 원금", loan_principal, ""],
    ["연 이자율 (%)", annual_interest_rate, ""],
    ["대출 기간 (년)", loan_term_years, ""],
    ["[ 초기 투자 비용 ]", "", ""],
    ["초기 현금 투자", initial_cash_investment, "매매가 - 대출원금"],
    ["취득세 및 지방교육세", acquisition_tax, ""],
    ["매매 중개수수료 (매수)", brokerage_fee_buy, ""],
    ["인테리어 비용", interior_cost, ""],
    ["[ 연간 보유 비용 ]", "", ""],
    ["재산세 (년)", property_tax_per_year, ""],
    ["[ 매도 시 비용 ]", "", ""],
    ["중개수수료율 (%)", brokerage_fee_rate, "부가세 포함"],
]
df_input = pd.DataFrame(data_input_data[1:], columns=data_input_data[0])


# --- Sheet 2: 월별 상환 계산 (원금균등 방식) ---
monthly_interest_rate = annual_interest_rate / 100 / 12
total_payments = loan_term_years * 12
monthly_principal_payment = loan_principal / total_payments

repayment_schedule = []
remaining_balance = loan_principal
repayment_schedule.append([0, 0, 0, 0, loan_principal])

for month in range(1, total_payments + 1):
    interest_payment = remaining_balance * monthly_interest_rate
    total_monthly_payment = monthly_principal_payment + interest_payment
    remaining_balance -= monthly_principal_payment
    repayment_schedule.append([month, total_monthly_payment, interest_payment, monthly_principal_payment, remaining_balance])

df_repayment = pd.DataFrame(repayment_schedule, columns=["월차", "월 상환액", "월 이자", "월 원금", "잔여 원금"])


# --- Sheet 3: 손익 분석 및 시나리오 ---
# A. 현재까지의 재무 상태 분석
current_loan_balance = df_repayment.loc[df_repayment['월차'] == elapsed_months, '잔여 원금'].iloc[0]
total_interest_paid_to_date = df_repayment.loc[df_repayment['월차'] <= elapsed_months, '월 이자'].sum()

# 재산세 계산 (1년에 2번 납부: 상반기, 하반기)
years_elapsed = elapsed_months / 12
property_tax_paid_to_date = property_tax_per_year * years_elapsed

total_investment_to_date = initial_investment_costs + total_interest_paid_to_date + property_tax_paid_to_date
current_net_assets = current_price - current_loan_balance
current_total_loss = total_investment_to_date - current_net_assets

# B. 5년 후 미래 시나리오 분석
future_simulation_end_month = elapsed_months + scenario_months
future_loan_balance = df_repayment.loc[df_repayment['월차'] == future_simulation_end_month, '잔여 원금'].iloc[0]
future_interest_paid = df_repayment.loc[(df_repayment['월차'] > elapsed_months) & (df_repayment['월차'] <= future_simulation_end_month), '월 이자'].sum()

# 미래 재산세 (1년에 2번)
scenario_years = scenario_months / 12
future_property_tax = property_tax_per_year * scenario_years

future_holding_cost = future_interest_paid + future_property_tax

# 매도 중개수수료 (동적 계산 - 매도가 기준)
selling_brokerage_fee_rate = get_brokerage_rate(scenario_sale_price)
selling_brokerage_fee = scenario_sale_price * (selling_brokerage_fee_rate / 100)

net_proceeds_after_sale = scenario_sale_price - future_loan_balance - selling_brokerage_fee
final_pnl = net_proceeds_after_sale - (total_investment_to_date + future_holding_cost)

# --- 투자 원금 분해 ---
total_cash_invested = initial_cash_investment
total_taxes_paid = acquisition_tax
total_misc_costs = brokerage_fee_buy + interior_cost
initial_investment_costs_updated = total_cash_invested + total_taxes_paid + total_misc_costs

# --- 분석 데이터 ---
analysis_data = [
    ["구분", "값", "비고"],
    ["[ 초기 투자 원금 ]", "", ""],
    ["초기 현금 투자", total_cash_invested, "매매가 - 대출원금"],
    ["[ 초기 기타 비용 ]", "", ""],
    ["취득세 및 지방교육세", total_taxes_paid, ""],
    ["매매 중개수수료 (매수)", brokerage_fee_buy, f"{brokerage_fee_rate}% 적용"],
    ["인테리어 비용", total_misc_costs - brokerage_fee_buy, ""],
    ["[ 현재 재무 상태 분석 ]", "", ""],
    ["총 보유 기간 (개월)", elapsed_months, "매매일부터 현재까지"],
    ["현재까지 낸 이자 (누적)", total_interest_paid_to_date, ""],
    ["현재까지 낸 재산세 (누적)", property_tax_paid_to_date, "1년에 2번 납부"],
    ["총 투자 원금 (과거 누적)", total_investment_to_date, "초기비용+이자+재산세 합계"],
    ["현재 시세", current_price, ""],
    ["현재 대출 잔액", current_loan_balance, f"{elapsed_months}개월차 상환 후"],
    ["현재 순자산 가치", current_net_assets, "(현재 시세 - 대출 잔액)"],
    ["현재 총 손실", current_total_loss, "(총 투자 원금 - 순자산)"],
    ["", "", ""],
    ["[ 5년 후 미래 시나리오 분석 ]", "", ""],
    ["분석 기간 (개월)", scenario_months, f"현재로부터 {scenario_months//12}년 후"],
    ["예상 매도가", scenario_sale_price, "⭐ 이 값만 바꾸면 나머지가 자동 계산됩니다!"],
    ["", "", ""],
    ["추가 이자 비용", future_interest_paid, f"{scenario_months//12}년간"],
    ["추가 재산세", future_property_tax, f"{scenario_months//12}년간 (1년 2회)"],
    ["추가 보유 비용 (이자+재산세)", future_holding_cost, f"{scenario_months//12}년간 합계"],
    ["해당 시점 대출 잔액", future_loan_balance, f"{future_simulation_end_month}개월차"],
    ["매도 중개수수료", selling_brokerage_fee, f"{selling_brokerage_fee_rate}% 적용"],
    ["매도 후 순수익", net_proceeds_after_sale, "(매도가 - 대출 - 수수료)"],
    ["", "", ""],
    ["최종 손익", final_pnl, "(순수익 - 총 지출액)"],
]
df_analysis = pd.DataFrame(analysis_data[1:], columns=analysis_data[0])


# ===================================================================
# 3. 엑셀 파일 생성 (인터랙티브 기능 추가)
# ===================================================================
file_name = f"budongsan_analysis_{current_date.strftime('%Y%m%d')}.xlsx"

with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    df_input.to_excel(writer, sheet_name='데이터 입력', index=False)
    df_repayment.to_excel(writer, sheet_name='월별 상환 계산', index=False)
    df_analysis.to_excel(writer, sheet_name='손익 분석 및 시나리오', index=False)

    workbook = writer.book
    sheet_analysis = workbook['손익 분석 및 시나리오']

    # --- 자동화 수식 (엑셀에서 예상 매도가만 변경하면 나머지 자동 계산) ---
    # B25: 해당 시점 대출 잔액 (과거 기간 + 미래 기간)
    # OFFSET 함수 사용: 월별상환 시트 E1부터 시작, 행 수 = B9+B19, 열 0
    sheet_analysis['B25'] = "=OFFSET('월별 상환 계산'!$E$1, B9+B19, 0)"

    # B26: 매도 중개수수료 (동적)
    # 5억 미만: 0.4%, 5억 이상: 0.3% 적용
    sheet_analysis['B26'] = "=IF(B20<500000000, B20*0.004, B20*0.003)"

    # B27: 매도 후 순수익
    sheet_analysis['B27'] = "=B20-B25-B26"

    # B29: 최종 손익 (순수익 - (총 투자원금 + 추가 보유 비용))
    sheet_analysis['B29'] = "=B27-(B12+B24)"

print(f"[OK] 최종 수정된 엑셀 파일 '{file_name}' 생성이 완료되었습니다.")
print(f"[Path] 파일 위치: {os.path.abspath(file_name)}")
print("[Info] 이제 엑셀 파일의 '손익 분석' 시트에서 '예상 매도가'(B14)만 바꾸면 나머지 값들이 실시간으로 변경됩니다!")
