<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<!--#include file="h_gamble.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
dim bet
dim over
dim pages
pages="gamble.asp"
bet=trim(Request.Form("Bet"))
select case Request.Form("Action")
case "开局"
over=0
call GameStart(bet)
case "再来一局"
over=0
call GameStart(bet)
case "要牌"
over=0
call GiveUserPoker()
if CanAddPoker() then
GiveAntiPoker()
end if
case "开牌"
if CanAddPoker() then
over=0
GiveAntiPoker()
else
''''''''''''''''''''''''
Session("n_Result")=Result()	
Session("n_Begin")=false
			
dim value	
over=0	
value=CLng(Session("n_Bet"))
Session("n_TotalMoney")=Session("n_TotalMoney")+value
'value=CCur(Session("n_Bet"))
select case CInt(Session("n_Result"))
case 1 'win
				
value=value
case 2 'lost
value=0-value
case 3
value=0
case else
value=0
end select
conn.execute("Update 用户 set 银两=银两+"&value&",赌次数=赌次数+1,魅力=魅力-1,赢钱=赢钱+"&value&" where 姓名='"&aqjh_name&"'") 				
Session("n_TotalMoney")=Session("n_TotalMoney")+value		
Session("n_Bet")=0

end if
case "不玩了"
pages="../jh.asp"
case "返回"
'在这里初始化数据
'	Session("n_Init")=false
'	Session("n_TotalMoney")=1000
'case "不玩了"
'退出吧。
pages="../jh.asp"
case else
call GameStart(bet)
end select


if over=0 then
Response.Redirect pages
end if
	
%>
<html><script language="JavaScript">                                                                  </script></html>
<html></html>