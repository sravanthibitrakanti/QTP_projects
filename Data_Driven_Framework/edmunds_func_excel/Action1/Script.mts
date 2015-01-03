Dim  mon_aff1,down_pay,apr,mon_pay,zip_code
Dim after_trade,loan_term,trade,vurl,val,vzip_code,vapr,vmon_pay
Dim vtrade_in,v_owed,vdownpay,vdown
Dim total_downpay
Dim testing,vtest
Dim i,k
Dim found,ss
Dim xLpath,xLSheet,xlrow,xlcol
Dim vlread
xLsheet="TestData"
xLpath="C:\Program Files\HP\QuickTest Professional\Tests\edmunds_xl"
vurl="www.edmunds.com/calculators"



For i=2 to 4
mon_aff1=xlread_cell(xLpath,xLsheet, i,1)
print mon_aff1
zip_code=xlread_cell(xLpath,xLsheet,i,2)
print zip_code
down_pay=xlread_cell(xLpath,xLsheet,i,3)
print down_pay
apr=xlread_cell(xLpath,xLsheet,i,4)
print apr
mon_pay=xlread_cell(xLpath,xLsheet,i,5)
print mon_pay
trade=xlread_cell(xLpath,xLsheet,i,6)
print trade
loan_term=xlread_cell(xLpath,xLsheet,i,7)
print loan_term
after_trade=xlread_cell(xLpath,xLsheet,i,8)
print after_trade

step0(vurl)
step1(mon_aff1)
step2()
step3()
step4()

Next



Function step0(furl)
 print "starting step0"
 systemutil.run"iexplore.exe",furl
End Function

Function step1(mon_aff2)
print "starting step1"
browser("Calculator: How Much Car").page("page1").WebEdit("monthly_affordability").Set mon_aff2
browser("Calculator: How Much Car").page("page1").WebElement("Go").Click
wait 3
End Function

Function step2()
  print "starting step2"
browser("Calculator: How Much Car").page("page2").WebEdit("ac_zip_code").Set zip_code
browser("Calculator: How Much Car").page("page2").WebEdit("ac_cash_down_payment").Set down_pay
browser("Calculator: How Much Car").page("page2").WebEdit("ac_market_finance_rate").Set apr
browser("Calculator: How Much Car").page("page2").WebEdit("ac_monthly_payment").Set mon_pay
vmon_pay=browser("Calculator: How Much Car").Page("page2").WebEdit("ac_monthly_payment").GetROProperty("value")
if cdbl(vmon_pay)=mon_pay Then
	print vmon_pay
	print mon_pay
	print true
	else
   print vmon_pay
	print mon_pay
	print false
End if
vzip_code=browser("Calculator: How Much Car").page("page2").WebEdit("ac_zip_code").GetROProperty("value")
print vzip_code
if cdbl(vzip_code)=zip_code Then
	print True
	else
	print False
End if

browser("Calculator: How Much Car").page("page2").WebEdit("am_owed_after_trade_in").Set after_trade
browser("Calculator: How Much Car").page("page2").WebList("loan_term").Select loan_term
browser("Calculator: How Much Car").page("page2").WebEdit("value_trade_in").Set trade
browser("Calculator: How Much Car").page("page2").WebElement("Calculate").click
End Function


Function step3()
print "starting step3"
total_downpay=(cint(down_pay)+cint(trade))-(cint(after_trade))
total_downpay="$"+cstr(total_downpay)
print total_downpay
vdown=browser("Calculator: How Much Car").Page("page2").WebElement("downpayment").GetROProperty("innertext")
k=Replace(vdown,",","")
xlwrite_cell xLpath,xLsheet,i,9,k
print  k
if k=total_downpay Then
	print "pass "
	xlwrite_cell xLpath,xLsheet,i,10,"pass"
	'print vdown
	'print total_downpay
	else
	print "Fail"
	xlwrite_cell xLpath,xLsheet,i,10,"Fail"
	'print k
	'print total_downpay
End if
End Function

Function step4()
   		print "Last Step - close browser"
		'TO DO :  Handle any pop-ups that come up.
		' STEP N : Close the browser
		Browser("Calculator: How Much Car").Close
End Function







