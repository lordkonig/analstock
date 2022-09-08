from flask import Flask, render_template, request, jsonify #, requests
# from vsearch import search4letters
from creon import Creon
import time, calendar
# import constants

# from setuptools import setup, find_packages
# setup(
#     name='flaskr',
#     packages=find_packages(),
#     include_package_data=True,
#     install_requires=[
#         'flask',
#     ],
#     setup_requires=[
#         'pytest-runner',
#     ],
#     tests_require=[
#         'pytest',
#     ],
# )





app = Flask(__name__)
c = Creon()

@app.route('/connection', methods=['GET', 'POST', 'DELETE'])
def handle_connect():
    c = Creon()
    if request.method == 'GET':
        # check connection status
        return jsonify(c.connected())
    elif request.method == 'POST':
        # make connection
        data = request.get_json()
        # _id = data['id']
        # _pwd = data['pwd']
        # _pwdcert = data['pwdcert']
        _id = data['insooya']
        _pwd = data['Passwo1!']
        _pwdcert = data['Password12!']
        return jsonify(c.connect(_id, _pwd, _pwdcert))
    elif request.method == 'DELETE':
        # disconnect
        res = c.disconnect()
        c.kill_client()
        return jsonify(res)





# 출력 템플릿 함수를 저장해놓고 재사용한다.
def template(content):
    return f'''
        <html>
        <head>

        </head>
        <body>
            {content}
        </body>
        </html>
        '''

# @app.route('/', methods=['GET', 'POST'])
@app.route('/', methods=['POST'])
def home() -> str :

    c.connect('insooya', 'Passwo1!', 'Password12!')

    title1 = '주식종목 재무제표분석!'
    content = '''     '''
    for code1 in [ 'A005010' ]:
        code_high5morename1 = c.instCpCodeMgr.CodeToName(code1)
        if c.get_todayclose(code1) > c.get_lastclose(code1):
            if ( c.MktCapital_bps(code1) > c.get_todayclose(code1) ) and ( 10*c.MktCapital_eps(code1) > c.get_todayclose(code1) ) and ( c.MktCapital_epsroe(code1) > c.get_todayclose(code1) ) :
                return render_template('results.html', title0 = title1, ) + str( template(content) ) + str('<h2>') + str(code_high5morename1) + str(' (') + str(code1) + str(', ') + str(c.instCpCodeMgr.CodeToName( c.instCpCodeMgr.GetStockIndustryCode(code1) )) + str(')') + str(' +') + str( round( ( (c.get_todayclose(code1))/(c.get_lastclose(code1)) -1), 3)*100 ) + str('%★') + \
                    str(',</h2>   <h3>시가총액 ') + str( round( c.MktCapital(code1)/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>★순유동자산(★시가총액보다크면 저평가된종목!) ') + str( round( 10000*c.MktCapital_initialmoney(code1)*( c.MktCapital_moneyrate(code1) - c.MktCapital_debt(code1) )/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>■매출액증가율 ') + str( round(c.MktCapital_salesriserate1(code1),1) ) + str('%,   □마지막분기매출액증가율 ') + str( round(c.MktCapital_salesriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■순이익증가율 ') + str( round(c.MktCapital_prftriserate1(code1),1) ) + str('%,   □마지막분기순이익증가율 ') + str( round(c.MktCapital_prftriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■매출대비영업이익률 ') + str( round(c.MktCapital_salesprft1(code1),1) ) + str('%,   □마지막분기매출대비영업이익률 ') + str( round(c.MktCapital_salesprft2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■영업이익증가율 ') + str( round(c.MktCapital_salesprftrise1(code1),1) ) + str('%,   □마지막분기영업이익증가율 ') + str( round(c.MktCapital_salesprftrise2(code1),1) ) + str('%,   </h3>') + \
                    str('최근1주거래대금 ') + str( round( (c.get_todayTrMoney(code1) +c.get_1daybef_TrMoney(code1) +c.get_2daybef_TrMoney(code1) +c.get_3daybef_TrMoney(code1) +c.get_4daybef_TrMoney(code1) )/100000000,1) ) + str('억') + \
                    str(',   <h3>☞부채비율 ') + str( c.MktCapital_debt(code1) ) + str('%,   ') + str('☞유보율 ') + str( c.MktCapital_moneyrate(code1) ) + str('%,   ') + str('☞ROE(이익률) ') + str( c.MktCapital_roe(code1) ) + str('%,   ') + \
                    str('☞마지막분기ROE ') + str( c.MktCapital_roe2(code1) ) + str('%,   </h3>') + \
                    str('<h3>◈◈부채,유보,ROE통합지수 ') + str( round( ( ( c.MktCapital_moneyrate(code1) * c.MktCapital_roe(code1) ) / ( 1 + c.MktCapital_debt(code1) ) ) ,1) ) + str(',   </h3>') + \
                    str('★BPS,EPS,ROE기준 특히저평가(+),   ') + \
                    str('베타계수 ') + str( round(c.MktCapital_beta1(code1),2) ) + str(',   ') + \
                    str('<h3>PER ') + str( c.MktCapital_per(code1) ) + str(',   ') + str('★pbr ') + str( c.MktCapital_pbr(code1) ) + str(',</h3>   ') + str('현재주가 ') + str(c.get_todayclose(code1)) + str('원')
            else:
                return render_template('results.html', title0 = title1, ) + str( template(content) ) + str('<h2>') + str(code_high5morename1) + str(' (') + str(code1) + str(', ') + str(c.instCpCodeMgr.CodeToName( c.instCpCodeMgr.GetStockIndustryCode(code1) )) + str(')') + str(' +') + str( round( ( (c.get_todayclose(code1))/(c.get_lastclose(code1)) -1), 3)*100 ) + str('%★') + \
                    str(',</h2>   <h3>시가총액 ') + str( round( c.MktCapital(code1)/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>★순유동자산(★시가총액보다크면 저평가된종목!) ') + str( round( 10000*c.MktCapital_initialmoney(code1)*( c.MktCapital_moneyrate(code1) - c.MktCapital_debt(code1) )/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>■매출액증가율 ') + str( round(c.MktCapital_salesriserate1(code1),1) ) + str('%,   □마지막분기매출액증가율 ') + str( round(c.MktCapital_salesriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■순이익증가율 ') + str( round(c.MktCapital_prftriserate1(code1),1) ) + str('%,   □마지막분기순이익증가율 ') + str( round(c.MktCapital_prftriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■매출대비영업이익률 ') + str( round(c.MktCapital_salesprft1(code1),1) ) + str('%,   □마지막분기매출대비영업이익률 ') + str( round(c.MktCapital_salesprft2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■영업이익증가율 ') + str( round(c.MktCapital_salesprftrise1(code1),1) ) + str('%,   □마지막분기영업이익증가율 ') + str( round(c.MktCapital_salesprftrise2(code1),1) ) + str('%,   </h3>') + \
                    str('최근1주거래대금 ') + str( round( (c.get_todayTrMoney(code1) +c.get_1daybef_TrMoney(code1) +c.get_2daybef_TrMoney(code1) +c.get_3daybef_TrMoney(code1) +c.get_4daybef_TrMoney(code1) )/100000000,1) ) + str('억') + \
                    str(',   <h3>☞부채비율 ') + str( c.MktCapital_debt(code1) ) + str('%,   ') + str('☞유보율 ') + str( c.MktCapital_moneyrate(code1) ) + str('%,   ') + str('☞ROE(이익률) ') + str( c.MktCapital_roe(code1) ) + str('%,   ') + \
                    str('☞마지막분기ROE ') + str( c.MktCapital_roe2(code1) ) + str('%,   </h3>') + \
                    str('<h3>◈◈부채,유보,ROE통합지수 ') + str( round( ( ( c.MktCapital_moneyrate(code1) * c.MktCapital_roe(code1) ) / ( 1 + c.MktCapital_debt(code1) ) ) ,1) ) + str(',   </h3>') + \
                    str('★BPS,EPS,ROE기준 특히저평가XX,   ') + \
                    str('베타계수 ') + str( round(c.MktCapital_beta1(code1),2) ) + str(',   ') + \
                    str('<h3>PER ') + str( c.MktCapital_per(code1) ) + str(',   ') + str('★pbr ') + str( c.MktCapital_pbr(code1) ) + str(',</h3>   ') + str('현재주가 ') + str(c.get_todayclose(code1)) + str('원')
        else:
            if ( c.MktCapital_bps(code1) > c.get_todayclose(code1) ) and ( 10*c.MktCapital_eps(code1) > c.get_todayclose(code1) ) and ( c.MktCapital_epsroe(code1) > c.get_todayclose(code1) ) :
                return render_template('results.html', title0 = title1, ) + str( template(content) ) + str('<h2>') + str(code_high5morename1) + str(' (') + str(code1) + str(', ') + str(c.instCpCodeMgr.CodeToName( c.instCpCodeMgr.GetStockIndustryCode(code1) )) + str(')') + str(' ') + str( round( ( (c.get_todayclose(code1))/(c.get_lastclose(code1)) -1), 3)*100 ) + str('%★') + \
                    str(',</h2>   <h3>시가총액 ') + str( round( c.MktCapital(code1)/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>★순유동자산(★시가총액보다크면 저평가된종목!) ') + str( round( 10000*c.MktCapital_initialmoney(code1)*( c.MktCapital_moneyrate(code1) - c.MktCapital_debt(code1) )/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>■매출액증가율 ') + str( round(c.MktCapital_salesriserate1(code1),1) ) + str('%,   □마지막분기매출액증가율 ') + str( round(c.MktCapital_salesriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■순이익증가율 ') + str( round(c.MktCapital_prftriserate1(code1),1) ) + str('%,   □마지막분기순이익증가율 ') + str( round(c.MktCapital_prftriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■매출대비영업이익률 ') + str( round(c.MktCapital_salesprft1(code1),1) ) + str('%,   □마지막분기매출대비영업이익률 ') + str( round(c.MktCapital_salesprft2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■영업이익증가율 ') + str( round(c.MktCapital_salesprftrise1(code1),1) ) + str('%,   □마지막분기영업이익증가율 ') + str( round(c.MktCapital_salesprftrise2(code1),1) ) + str('%,   </h3>') + \
                    str('최근1주거래대금 ') + str( round( (c.get_todayTrMoney(code1) +c.get_1daybef_TrMoney(code1) +c.get_2daybef_TrMoney(code1) +c.get_3daybef_TrMoney(code1) +c.get_4daybef_TrMoney(code1) )/100000000,1) ) + str('억') + \
                    str(',   <h3>☞부채비율 ') + str( c.MktCapital_debt(code1) ) + str('%,   ') + str('☞유보율 ') + str( c.MktCapital_moneyrate(code1) ) + str('%,   ') + str('☞ROE(이익률) ') + str( c.MktCapital_roe(code1) ) + str('%,   ') + \
                    str('☞마지막분기ROE ') + str( c.MktCapital_roe2(code1) ) + str('%,   </h3>') + \
                    str('<h3>◈◈부채,유보,ROE통합지수 ') + str( round( ( ( c.MktCapital_moneyrate(code1) * c.MktCapital_roe(code1) ) / ( 1 + c.MktCapital_debt(code1) ) ) ,1) ) + str(',   </h3>') + \
                    str('★BPS,EPS,ROE기준 특히저평가(+),   ') + \
                    str('베타계수 ') + str( round(c.MktCapital_beta1(code1),2) ) + str(',   ') + \
                    str('<h3>PER ') + str( c.MktCapital_per(code1) ) + str(',   ') + str('★pbr ') + str( c.MktCapital_pbr(code1) ) + str(',</h3>   ') + str('현재주가 ') + str(c.get_todayclose(code1)) + str('원')
            else:
                return render_template('results.html', title0 = title1, ) + str( template(content) ) + str('<h2>') + str(code_high5morename1) + str(' (') + str(code1) + str(', ') + str(c.instCpCodeMgr.CodeToName( c.instCpCodeMgr.GetStockIndustryCode(code1) )) + str(')') + str(' ') + str( round( ( (c.get_todayclose(code1))/(c.get_lastclose(code1)) -1), 3)*100 ) + str('%★') + \
                    str(',</h2>   <h3>시가총액 ') + str( round( c.MktCapital(code1)/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>★순유동자산(★시가총액보다크면 저평가된종목!) ') + str( round( 10000*c.MktCapital_initialmoney(code1)*( c.MktCapital_moneyrate(code1) - c.MktCapital_debt(code1) )/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>■매출액증가율 ') + str( round(c.MktCapital_salesriserate1(code1),1) ) + str('%,   □마지막분기매출액증가율 ') + str( round(c.MktCapital_salesriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■순이익증가율 ') + str( round(c.MktCapital_prftriserate1(code1),1) ) + str('%,   □마지막분기순이익증가율 ') + str( round(c.MktCapital_prftriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■매출대비영업이익률 ') + str( round(c.MktCapital_salesprft1(code1),1) ) + str('%,   □마지막분기매출대비영업이익률 ') + str( round(c.MktCapital_salesprft2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■영업이익증가율 ') + str( round(c.MktCapital_salesprftrise1(code1),1) ) + str('%,   □마지막분기영업이익증가율 ') + str( round(c.MktCapital_salesprftrise2(code1),1) ) + str('%,   </h3>') + \
                    str('최근1주거래대금 ') + str( round( (c.get_todayTrMoney(code1) +c.get_1daybef_TrMoney(code1) +c.get_2daybef_TrMoney(code1) +c.get_3daybef_TrMoney(code1) +c.get_4daybef_TrMoney(code1) )/100000000,1) ) + str('억') + \
                    str(',   <h3>☞부채비율 ') + str( c.MktCapital_debt(code1) ) + str('%,   ') + str('☞유보율 ') + str( c.MktCapital_moneyrate(code1) ) + str('%,   ') + str('☞ROE(이익률) ') + str( c.MktCapital_roe(code1) ) + str('%,   ') + \
                    str('☞마지막분기ROE ') + str( c.MktCapital_roe2(code1) ) + str('%,   </h3>') + \
                    str('<h3>◈◈부채,유보,ROE통합지수 ') + str( round( ( ( c.MktCapital_moneyrate(code1) * c.MktCapital_roe(code1) ) / ( 1 + c.MktCapital_debt(code1) ) ) ,1) ) + str(',   </h3>') + \
                    str('★BPS,EPS,ROE기준 특히저평가XX,   ') + \
                    str('베타계수 ') + str( round(c.MktCapital_beta1(code1),2) ) + str(',   ') + \
                    str('<h3>PER ') + str( c.MktCapital_per(code1) ) + str(',   ') + str('★pbr ') + str( c.MktCapital_pbr(code1) ) + str(',</h3>   ') + str('현재주가 ') + str(c.get_todayclose(code1)) + str('원')







# @app.route('/stockanal', methods=['GET', 'POST'])
@app.route('/stockanal', methods=['POST'])
def analyze() -> str:

    c.connect('insooya', 'Passwo1!', 'Password12!')

    title1 = '주식종목 재무제표분석!'
    code2 = request.form['code3']
    content = '''     '''
    for code1 in [ code2 ]:
        code_high5morename1 = c.instCpCodeMgr.CodeToName(code1)
        if c.get_todayclose(code1) > c.get_lastclose(code1):
            if ( c.MktCapital_bps(code1) > c.get_todayclose(code1) ) and ( 10*c.MktCapital_eps(code1) > c.get_todayclose(code1) ) and ( c.MktCapital_epsroe(code1) > c.get_todayclose(code1) ) :
                return render_template('results.html', title0 = title1, ) + str( template(content) ) + str('<h2>') + str(code_high5morename1) + str(' (') + str(code1) + str(', ') + str(c.instCpCodeMgr.CodeToName( c.instCpCodeMgr.GetStockIndustryCode(code1) )) + str(')') + str(' +') + str( round( ( (c.get_todayclose(code1))/(c.get_lastclose(code1)) -1), 3)*100 ) + str('%★') + \
                    str(',</h2>   <h3>시가총액 ') + str( round( c.MktCapital(code1)/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>★순유동자산(★시가총액보다크면 저평가된종목!) ') + str( round( 10000*c.MktCapital_initialmoney(code1)*( c.MktCapital_moneyrate(code1) - c.MktCapital_debt(code1) )/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>■매출액증가율 ') + str( round(c.MktCapital_salesriserate1(code1),1) ) + str('%,   □마지막분기매출액증가율 ') + str( round(c.MktCapital_salesriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■순이익증가율 ') + str( round(c.MktCapital_prftriserate1(code1),1) ) + str('%,   □마지막분기순이익증가율 ') + str( round(c.MktCapital_prftriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■매출대비영업이익률 ') + str( round(c.MktCapital_salesprft1(code1),1) ) + str('%,   □마지막분기매출대비영업이익률 ') + str( round(c.MktCapital_salesprft2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■영업이익증가율 ') + str( round(c.MktCapital_salesprftrise1(code1),1) ) + str('%,   □마지막분기영업이익증가율 ') + str( round(c.MktCapital_salesprftrise2(code1),1) ) + str('%,   </h3>') + \
                    str('최근1주거래대금 ') + str( round( (c.get_todayTrMoney(code1) +c.get_1daybef_TrMoney(code1) +c.get_2daybef_TrMoney(code1) +c.get_3daybef_TrMoney(code1) +c.get_4daybef_TrMoney(code1) )/100000000,1) ) + str('억') + \
                    str(',   <h3>☞부채비율 ') + str( c.MktCapital_debt(code1) ) + str('%,   ') + str('☞유보율 ') + str( c.MktCapital_moneyrate(code1) ) + str('%,   ') + str('☞ROE(이익률) ') + str( c.MktCapital_roe(code1) ) + str('%,   ') + \
                    str('☞마지막분기ROE ') + str( c.MktCapital_roe2(code1) ) + str('%,   </h3>') + \
                    str('<h3>◈◈부채,유보,ROE통합지수 ') + str( round( ( ( c.MktCapital_moneyrate(code1) * c.MktCapital_roe(code1) ) / ( 1 + c.MktCapital_debt(code1) ) ) ,1) ) + str(',   </h3>') + \
                    str('★BPS,EPS,ROE기준 특히저평가(+),   ') + \
                    str('베타계수 ') + str( round(c.MktCapital_beta1(code1),2) ) + str(',   ') + \
                    str('<h3>PER ') + str( c.MktCapital_per(code1) ) + str(',   ') + str('★pbr ') + str( c.MktCapital_pbr(code1) ) + str(',</h3>   ') + str('현재주가 ') + str(c.get_todayclose(code1)) + str('원')
            else:
                return render_template('results.html', title0 = title1, ) + str( template(content) ) + str('<h2>') + str(code_high5morename1) + str(' (') + str(code1) + str(', ') + str(c.instCpCodeMgr.CodeToName( c.instCpCodeMgr.GetStockIndustryCode(code1) )) + str(')') + str(' +') + str( round( ( (c.get_todayclose(code1))/(c.get_lastclose(code1)) -1), 3)*100 ) + str('%★') + \
                    str(',</h2>   <h3>시가총액 ') + str( round( c.MktCapital(code1)/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>★순유동자산(★시가총액보다크면 저평가된종목!) ') + str( round( 10000*c.MktCapital_initialmoney(code1)*( c.MktCapital_moneyrate(code1) - c.MktCapital_debt(code1) )/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>■매출액증가율 ') + str( round(c.MktCapital_salesriserate1(code1),1) ) + str('%,   □마지막분기매출액증가율 ') + str( round(c.MktCapital_salesriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■순이익증가율 ') + str( round(c.MktCapital_prftriserate1(code1),1) ) + str('%,   □마지막분기순이익증가율 ') + str( round(c.MktCapital_prftriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■매출대비영업이익률 ') + str( round(c.MktCapital_salesprft1(code1),1) ) + str('%,   □마지막분기매출대비영업이익률 ') + str( round(c.MktCapital_salesprft2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■영업이익증가율 ') + str( round(c.MktCapital_salesprftrise1(code1),1) ) + str('%,   □마지막분기영업이익증가율 ') + str( round(c.MktCapital_salesprftrise2(code1),1) ) + str('%,   </h3>') + \
                    str('최근1주거래대금 ') + str( round( (c.get_todayTrMoney(code1) +c.get_1daybef_TrMoney(code1) +c.get_2daybef_TrMoney(code1) +c.get_3daybef_TrMoney(code1) +c.get_4daybef_TrMoney(code1) )/100000000,1) ) + str('억') + \
                    str(',   <h3>☞부채비율 ') + str( c.MktCapital_debt(code1) ) + str('%,   ') + str('☞유보율 ') + str( c.MktCapital_moneyrate(code1) ) + str('%,   ') + str('☞ROE(이익률) ') + str( c.MktCapital_roe(code1) ) + str('%,   ') + \
                    str('☞마지막분기ROE ') + str( c.MktCapital_roe2(code1) ) + str('%,   </h3>') + \
                    str('<h3>◈◈부채,유보,ROE통합지수 ') + str( round( ( ( c.MktCapital_moneyrate(code1) * c.MktCapital_roe(code1) ) / ( 1 + c.MktCapital_debt(code1) ) ) ,1) ) + str(',   </h3>') + \
                    str('★BPS,EPS,ROE기준 특히저평가XX,   ') + \
                    str('베타계수 ') + str( round(c.MktCapital_beta1(code1),2) ) + str(',   ') + \
                    str('<h3>PER ') + str( c.MktCapital_per(code1) ) + str(',   ') + str('★pbr ') + str( c.MktCapital_pbr(code1) ) + str(',</h3>   ') + str('현재주가 ') + str(c.get_todayclose(code1)) + str('원')
        else:
            if ( c.MktCapital_bps(code1) > c.get_todayclose(code1) ) and ( 10*c.MktCapital_eps(code1) > c.get_todayclose(code1) ) and ( c.MktCapital_epsroe(code1) > c.get_todayclose(code1) ) :
                return render_template('results.html', title0 = title1, ) + str( template(content) ) + str('<h2>') + str(code_high5morename1) + str(' (') + str(code1) + str(', ') + str(c.instCpCodeMgr.CodeToName( c.instCpCodeMgr.GetStockIndustryCode(code1) )) + str(')') + str(' ') + str( round( ( (c.get_todayclose(code1))/(c.get_lastclose(code1)) -1), 3)*100 ) + str('%★') + \
                    str(',</h2>   <h3>시가총액 ') + str( round( c.MktCapital(code1)/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>★순유동자산(★시가총액보다크면 저평가된종목!) ') + str( round( 10000*c.MktCapital_initialmoney(code1)*( c.MktCapital_moneyrate(code1) - c.MktCapital_debt(code1) )/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>■매출액증가율 ') + str( round(c.MktCapital_salesriserate1(code1),1) ) + str('%,   □마지막분기매출액증가율 ') + str( round(c.MktCapital_salesriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■순이익증가율 ') + str( round(c.MktCapital_prftriserate1(code1),1) ) + str('%,   □마지막분기순이익증가율 ') + str( round(c.MktCapital_prftriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■매출대비영업이익률 ') + str( round(c.MktCapital_salesprft1(code1),1) ) + str('%,   □마지막분기매출대비영업이익률 ') + str( round(c.MktCapital_salesprft2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■영업이익증가율 ') + str( round(c.MktCapital_salesprftrise1(code1),1) ) + str('%,   □마지막분기영업이익증가율 ') + str( round(c.MktCapital_salesprftrise2(code1),1) ) + str('%,   </h3>') + \
                    str('최근1주거래대금 ') + str( round( (c.get_todayTrMoney(code1) +c.get_1daybef_TrMoney(code1) +c.get_2daybef_TrMoney(code1) +c.get_3daybef_TrMoney(code1) +c.get_4daybef_TrMoney(code1) )/100000000,1) ) + str('억') + \
                    str(',   <h3>☞부채비율 ') + str( c.MktCapital_debt(code1) ) + str('%,   ') + str('☞유보율 ') + str( c.MktCapital_moneyrate(code1) ) + str('%,   ') + str('☞ROE(이익률) ') + str( c.MktCapital_roe(code1) ) + str('%,   ') + \
                    str('☞마지막분기ROE ') + str( c.MktCapital_roe2(code1) ) + str('%,   </h3>') + \
                    str('<h3>◈◈부채,유보,ROE통합지수 ') + str( round( ( ( c.MktCapital_moneyrate(code1) * c.MktCapital_roe(code1) ) / ( 1 + c.MktCapital_debt(code1) ) ) ,1) ) + str(',   </h3>') + \
                    str('★BPS,EPS,ROE기준 특히저평가(+),   ') + \
                    str('베타계수 ') + str( round(c.MktCapital_beta1(code1),2) ) + str(',   ') + \
                    str('<h3>PER ') + str( c.MktCapital_per(code1) ) + str(',   ') + str('★pbr ') + str( c.MktCapital_pbr(code1) ) + str(',</h3>   ') + str('현재주가 ') + str(c.get_todayclose(code1)) + str('원')
            else:
                return render_template('results.html', title0 = title1, ) + str( template(content) ) + str('<h2>') + str(code_high5morename1) + str(' (') + str(code1) + str(', ') + str(c.instCpCodeMgr.CodeToName( c.instCpCodeMgr.GetStockIndustryCode(code1) )) + str(')') + str(' ') + str( round( ( (c.get_todayclose(code1))/(c.get_lastclose(code1)) -1), 3)*100 ) + str('%★') + \
                    str(',</h2>   <h3>시가총액 ') + str( round( c.MktCapital(code1)/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>★순유동자산(★시가총액보다크면 저평가된종목!) ') + str( round( 10000*c.MktCapital_initialmoney(code1)*( c.MktCapital_moneyrate(code1) - c.MktCapital_debt(code1) )/100000000,1) ) + str('억,   </h3>') + \
                    str('<h3>■매출액증가율 ') + str( round(c.MktCapital_salesriserate1(code1),1) ) + str('%,   □마지막분기매출액증가율 ') + str( round(c.MktCapital_salesriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■순이익증가율 ') + str( round(c.MktCapital_prftriserate1(code1),1) ) + str('%,   □마지막분기순이익증가율 ') + str( round(c.MktCapital_prftriserate2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■매출대비영업이익률 ') + str( round(c.MktCapital_salesprft1(code1),1) ) + str('%,   □마지막분기매출대비영업이익률 ') + str( round(c.MktCapital_salesprft2(code1),1) ) + str('%,   </h3>') + \
                    str('<h3>■영업이익증가율 ') + str( round(c.MktCapital_salesprftrise1(code1),1) ) + str('%,   □마지막분기영업이익증가율 ') + str( round(c.MktCapital_salesprftrise2(code1),1) ) + str('%,   </h3>') + \
                    str('최근1주거래대금 ') + str( round( (c.get_todayTrMoney(code1) +c.get_1daybef_TrMoney(code1) +c.get_2daybef_TrMoney(code1) +c.get_3daybef_TrMoney(code1) +c.get_4daybef_TrMoney(code1) )/100000000,1) ) + str('억') + \
                    str(',   <h3>☞부채비율 ') + str( c.MktCapital_debt(code1) ) + str('%,   ') + str('☞유보율 ') + str( c.MktCapital_moneyrate(code1) ) + str('%,   ') + str('☞ROE(이익률) ') + str( c.MktCapital_roe(code1) ) + str('%,   ') + \
                    str('☞마지막분기ROE ') + str( c.MktCapital_roe2(code1) ) + str('%,   </h3>') + \
                    str('<h3>◈◈부채,유보,ROE통합지수 ') + str( round( ( ( c.MktCapital_moneyrate(code1) * c.MktCapital_roe(code1) ) / ( 1 + c.MktCapital_debt(code1) ) ) ,1) ) + str(',   </h3>') + \
                    str('★BPS,EPS,ROE기준 특히저평가XX,   ') + \
                    str('베타계수 ') + str( round(c.MktCapital_beta1(code1),2) ) + str(',   ') + \
                    str('<h3>PER ') + str( c.MktCapital_per(code1) ) + str(',   ') + str('★pbr ') + str( c.MktCapital_pbr(code1) ) + str(',</h3>   ') + str('현재주가 ') + str(c.get_todayclose(code1)) + str('원')





if __name__ == "__main__":
    app.run( debug=True, host="0.0.0.0", port="3333" )

