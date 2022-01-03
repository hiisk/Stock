# 주식에 관심이 생기면서 자동매매/주식가격예측 두 가지 프로젝트를 진행하였습니다.

## Auto Trade

# 기존 방식 
1. 관심이 있는 주식의 ETF 종목코드를 입력한다.
2. 9시부터 3시10분까지 변동성전략을 통해 구한 target_price의 값보다 높아질경우 매수를 진행한다.
3. 3시 10분이 지나면 전량 매도를 진행한다.

# 변경 방식
1. 매매량이 10000이 넘고 변동성 전략을 통한 3일, 2주 백테스팅 수익률이 1.003을 초과하고 ETF종목을 모두 리스트에 넣는다.

   (백테스팅시에 제일 높은 K값을 찾아서 설정)

2. 9시부터 3시10분까지 변동성전략을 통해 구한 target_price의 값보다 높아질경우 매수를 진행한다.
   
   (매수 성공인 경우 bought_list에 종목코드를 입력)

3. 2번에서 매수를 할 시점에 주식의 현재가가 target_price보다 1.002이상 높아졌을 경우 매수를 하지 않는다.
   
   (target_price < 현재가 < target_price< 1.002)

4. 매수가 완료된 종목에 한해 반 매도, 나머지 매도의 조건을 진행한다. ★등락이 심한 종목이라면 더욱 효율적★ 
   
   (반 매도: target_price * x <= 현재가, 나머지 매도: target_price * y <= 현재가)
   
   x = (target_tmp-1)/3 + 1 (1.01이하일 경우 그대로 입력)
   y = (target_tmp-1)*3 + 1 (1.01이하인 케이스까지 보완)

5. 3시 10분이 지나면 전량 매도를 진행한다.



# 중요사항
1. 거래수수료를 최소로 하기위해 ETF종목으로 진행(ETF수수료:0.019% 고려)
2. 변동성 전략 범위 설정(K값): 0.1부터 1사이에 0.01씩의 차이로 최적의 값을 도출하여 진행

# vscode로 작성하고 대신증권 creon으로 실행하였습니다.
위 코드는 "파이썬 증권 데이터 분석"을 기초로 제작되었습니다.(링크: https://bit.ly/30Yg38v)

관련영상 https://www.youtube.com/watch?v=5bTxyBeOVkA

# 변경사항
관련영상 유튜브에 있는 기존 버전과 다르게 저만의 방식을 추가하였습니다.

## Forecast

# 캡스톤 디자인 주제로 선정하여 구글 코랩으로 작성
# RNN LSTM GRU 방식 사용
# 2020년의 정보로 2021년을 예측하여 예측값과 실제값을 비교
