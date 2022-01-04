### [Auto Trade](https://github.com/hiisk/Stock/tree/main/Auto_Trade)

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
   
   (반 매도: target_price * 1.005 <= 현재가, 나머지 매도: target_price * x <= 현재가)

   (x = 2주의 백테스팅 수익률) --> 사실상 안팔릴 확률이 높습니다.

5. 3시 10분이 지나면 전량 매도를 진행한다.

### [Forecast](https://github.com/hiisk/Stock/tree/main/Forecast)

캡스톤 디자인 주제로 선정하여 구글 코랩으로 작성

RNN LSTM GRU 방식 사용

2020년의 정보로 2021년을 예측하여 예측값과 실제값을 비교
