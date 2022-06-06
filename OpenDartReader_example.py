import OpenDartReader
import pandas as pd
import time
from datetime import datetime
from pykrx import stock

api_key = '' # OpenDart API KEY
stock_names = ['003690'] #삼성전자 종목코드
dart = OpenDartReader(api_key)

# 정리된 Data. index 가 없으면 추가가 안되므로 dummy를 하나 넣어둔다.
df2 = pd.DataFrame(columns=['유동자산', '자산총계', '비유동자산', '부채총계', '자본총계', '매출액', '매출총이익', '영업이익',
                            '당기순이익', '영업활동현금흐름', '잉여현금흐름', '시가총액'], index=['1900-01-01']) 

# '11013'=1분기보고서, '11012' =반기보고서, '11014'=3분기보고서, '11011'=사업보고서
reprt_code = ['11013', '11012', '11014', '11011']

for stocks in stock_names:    
    fileName = f'C:/Users/USER/Desktop/PythonQuant/result_{str(stocks)}.xlsx'
    for i in range(2021, 2022): # OpenDart는 2015년부터 정보를 제공한다.
        # 더미 리스트 초기화
        current_assets = [0, 0, 0, 0] # 유동자산
        total_assets = [0, 0, 0, 0] #자산총계
        non_current_assets = [0, 0, 0, 0] #비유동자산
        liabilities = [0, 0, 0, 0] # 부채총계
        equity = [0, 0, 0, 0] # 자본총계
        revenue = [0, 0, 0, 0] # 매출액 
        grossProfit = [0, 0, 0, 0] # 매출총이익
        income = [0, 0, 0, 0] # 영업이익
        net_income = [0, 0, 0, 0] # 당기순이익
        cfo = [0, 0, 0, 0] # 영업활동현금흐름
        cfi = [0, 0, 0, 0] # 투자활동현금흐름
        fcf = [0, 0, 0, 0] # 잉여현금흐름 : 편의상 영업활동 - 투자활동 현금흐름으로 계산
        market_cap = [0, 0, 0, 0] # 시가총액
        date_year = str(i) # 년도 변수 지정
    
	    # '11013'=1분기보고서, '11012' =반기보고서, '11014'=3분기보고서, '11011'=사업보고서 순서대로 루프를 태운다.
        for j, k in enumerate(reprt_code): 
            df1 = pd.DataFrame() # Raw Data
            if str(type(dart.finstate_all(stocks, i, reprt_code=k, fs_div='CFS'))) == "<class 'NoneType'>":
                pass
            else: # 타입이 NoneType 이 아니면 읽어온다.
                df1 = df1.append(dart.finstate_all(stocks, i, reprt_code=k, fs_div='CFS')) 
                # 재무상태표 부분
                condition = (df1.sj_nm == '재무상태표') & (df1.account_nm == '유동자산') # 유동자산
                condition_0 = (df1.sj_nm == '재무상태표') & (df1.account_nm == '자산총계') # 자산총계
                condition_1 = (df1.sj_nm == '재무상태표') & (df1.account_nm == '비유동자산') # 비유동자산
                condition_2 = (df1.sj_nm == '재무상태표') & (df1.account_nm == '부채총계') # 부채총계
                condition_3 = (df1.sj_nm == '재무상태표') & \
                            ((df1.account_nm == '자본총계') | (df1.account_nm == '반기말자본') | (df1.account_nm == '3분기말자본') | (df1.account_nm == '분기말자본') | (df1.account_nm == '1분기말자본'))  #자본총계
                # 손익계산서 부분
                condition_4 = ((df1.sj_nm == '손익계산서') | (df1.sj_nm == '포괄손익계산서')) & ((df1.account_nm == '매출액') | (df1.account_nm == '수익(매출액)') | (df1.account_nm == '매출'))
                condition_5 = ((df1.sj_nm == '손익계산서') | (df1.sj_nm == '포괄손익계산서')) & ((df1.account_nm == '매출총이익') | (df1.account_nm == '매출총이익(손실)'))
                condition_6 = ((df1.sj_nm == '손익계산서') | (df1.sj_nm == '포괄손익계산서')) & \
                                ((df1.account_nm == '영업이익(손실)') | (df1.account_nm == '영업이익'))
                condition_7 = ((df1.sj_nm == '손익계산서') | (df1.sj_nm == '포괄손익계산서')) & \
                                ((df1.account_nm == '당기순이익(손실)') | (df1.account_nm == '당기순이익') | \
                                (df1.account_nm == '분기순이익') | (df1.account_nm == '분기순이익(손실)') | (df1.account_nm == '반기순이익') | (df1.account_nm == '반기순이익(손실)') | \
                                (df1.account_nm == '연결분기순이익') | (df1.account_nm == '연결반기순이익')| (df1.account_nm == '연결당기순이익')|(df1.account_nm == '연결분기(당기)순이익')|(df1.account_nm == '연결반기(당기)순이익')|\
                                (df1.account_nm == '연결분기순이익(손실)'))
                # 현금흐름표 부분
                condition_8 = (df1.sj_nm == '현금흐름표') & ((df1.account_nm == '영업활동으로 인한 현금흐름') | (df1.account_nm == '영업활동 현금흐름') | (df1.account_nm == '영업활동현금흐름') | (df1.account_nm == '영업활동으로 인한 순현금흐름'))
                condition_9 = (df1.sj_nm == '현금흐름표') & ((df1.account_nm == '투자활동으로 인한 현금흐름') | (df1.account_nm == '투자활동 현금흐름') | (df1.account_nm == '투자활동현금흐름') | (df1.account_nm == '투자활동으로 인한 순현금흐름'))
                            
                current_assets[j] = int(df1.loc[condition].iloc[0]['thstrm_amount'])
                total_assets[j] = int(df1.loc[condition_0].iloc[0]['thstrm_amount'])
                non_current_assets[j] = int(df1.loc[condition_1].iloc[0]['thstrm_amount'])
                liabilities[j] = int(df1.loc[condition_2].iloc[0]['thstrm_amount'])
                equity[j] = int(df1.loc[condition_3].iloc[0]['thstrm_amount'])
                revenue[j] = int(df1.loc[condition_4].iloc[0]['thstrm_amount'])
                grossProfit[j] = int(df1.loc[condition_5].iloc[0]['thstrm_amount'])
                income[j] = int(df1.loc[condition_6].iloc[0]['thstrm_amount'])
                net_income[j] = int(df1.loc[condition_7].iloc[0]['thstrm_amount'])
                cfo[j] = int(df1.loc[condition_8].iloc[0]['thstrm_amount'])
                cfi[j] = int(df1.loc[condition_9].iloc[0]['thstrm_amount'])
                fcf[j] = (cfo[j] - cfi[j])
                
                if k == '11013': # 1분기
                    date_month = '03'
                    date_day = 31 # 일만 계산할꺼니까 이것만 숫자로 지정

                elif k == '11012': # 2분기
                    date_month = '06'
                    date_day = 30

                elif k == '11014': # 3분기
                    date_month = '09'
                    date_day = 30

                else: # 4분기. 1 ~ 3분기 데이터를 더한다음 사업보고서에서 빼야 함
                    date_month = '12'
                    date_day = 30
                    revenue[j] = revenue[j] - (revenue[0] + revenue[1] + revenue[2])
                    grossProfit[j] = grossProfit[j] - (grossProfit[0] + grossProfit[1] + grossProfit[2])
                    income[j] = income[j] - (income[0] + income[1] + income[2])
                    net_income[j] = net_income[j] - (net_income[0] + net_income[1] + net_income[2])
                    fcf[j] = fcf[j] - (fcf[0] + fcf[1] + fcf[2])

                path_string = date_year + '-' + date_month + '-' + str(date_day)

                # 날짜 계산을 위한 변수 정의
                date = date_year + date_month + str(date_day)
                date_formated = datetime.strptime(date,"%Y%m%d") # datetime format 으로 변환
                
                # 주말에 대한 처리
                if date_formated.weekday() == 5:
                    date_day -= 1 # 토요일일 경우 1일을 뺀다
                elif date_formated.weekday() == 6:
                    date_day -= 2 # 일요일일 경우 2일을 뺀다

                # 3분기 추석에 대한 처리. 2020년은 1일을 빼고 2023년은 3일을 뺀다.
                if date_month == '09' and date_year == '2020':
                    date_day -= 1
                elif date_month == '09' and date_year == '2023':
                    date_day -= 3

                date = date_year + date_month + str(date_day) # 뺀 날짜에 대해 재정의
                date_2 = date_year + '-' + date_month + '-' + str(date_day) # 재정의된 날짜의 YYYY-MM-DD 형태
                df3 = stock.get_market_cap_by_date(date, date, str(stocks)) # 시가총액 데이터프레임 호출
                market_cap[j] = df3.loc[date_2]['시가총액'] # 시가총액 데이터 추출

                # 데이터프레임에 저장
                df2.loc[path_string] = [current_assets[j], total_assets[j], non_current_assets[j], liabilities[j], equity[j],
                                    revenue[j], grossProfit[j], income[j], net_income[j], cfo[j], fcf[j], market_cap[j]]                
                df2.tail()
            time.sleep(0.2) # 잦은 API 호출은 서버에서 IP 차단을 당하므로 호출 간격을 둔다.
    df2.drop(['1900-01-01'], inplace=True) # 첫 행 drop
    df2.to_excel(fileName) # 파일 저장. 저장할 때 파일의 경로지정을 해야 함. 각 종목코드별로 다른 이름으로 저장
    pd.set_option("display.max_columns", None)
    #display(df2)
    #df2 = pd.DataFrame(columns=['유동자산', '자산총계', '비유동자산', '부채총계', '자본총계', '매출액', '매출총이익', '영업이익',
    #                        '당기순이익', '영업활동현금흐름', '잉여현금흐름', '시가총액'], index=['1900-01-01'])
    #display(df2)
    
print(df2.shape)
display(df2)
length = len(df2)
print(length)
df4 = pd.DataFrame(columns=['PER','PBR','PSR','GP/A','POR','PCR','PFCR','NCAV/MK'], index=['1900-01-01'])
for i in range(3, length):

    #print(list(df2.index))
    #indexing = df2.iloc[i]['Unnamed: 0']
    
    
    indexing = list(df2.index)[i]
    PER = df2.iloc[i]['시가총액'] / (df2.iloc[i-3]['당기순이익'] + df2.iloc[i-2]['당기순이익'] + \
          df2.iloc[i-1]['당기순이익'] + df2.iloc[i]['당기순이익'])
    PBR = df2.iloc[i]['시가총액'] / (df2.iloc[i]['자산총계'] - df2.iloc[i]['부채총계'])
    PSR = df2.iloc[i]['시가총액'] / (df2.iloc[i-3]['매출액'] + df2.iloc[i-2]['매출액'] + \
          df2.iloc[i-1]['매출액'] + df2.iloc[i]['매출액'])
    GP_A = (df2.iloc[i-3]['매출총이익'] + df2.iloc[i-2]['매출총이익'] + \
          df2.iloc[i-1]['매출총이익'] + df2.iloc[i]['매출총이익']) / df2.iloc[i]['자산총계']
    POR = df2.iloc[i]['시가총액'] / (df2.iloc[i-3]['영업이익'] + df2.iloc[i-2]['영업이익'] + \
          df2.iloc[i-1]['영업이익'] + df2.iloc[i]['영업이익'])
    PCR = df2.iloc[i]['시가총액'] / (df2.iloc[i-3]['영업활동현금흐름'] + df2.iloc[i-2]['영업활동현금흐름'] + \
          df2.iloc[i-1]['영업활동현금흐름'] + df2.iloc[i]['영업활동현금흐름'])
    PFCR = df2.iloc[i]['시가총액'] / (df2.iloc[i-3]['잉여현금흐름'] + df2.iloc[i-2]['잉여현금흐름'] + \
          df2.iloc[i-1]['잉여현금흐름'] + df2.iloc[i]['잉여현금흐름'])
    NCAV_MK = (df2.iloc[i]['유동자산'] - df2.iloc[i]['부채총계']) / df2.iloc[i]['시가총액']

    df4.loc[indexing] = [PER, PBR, PSR, GP_A, POR, PCR, PFCR, NCAV_MK]

df4.drop(['1900-01-01'], inplace=True) # 첫 행 drop
pd.set_option("display.max_columns", None)
display(df4)    
    
