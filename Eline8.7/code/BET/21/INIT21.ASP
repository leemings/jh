<%@ LANGUAGE=VBScript codepage ="936" %>
<%

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="Select  * from 用户 where 姓名='"&sjjh_name&"'"
set rs=conn.Execute(sql)
nowmeili=rs("魅力")
if nowmeili<10 then
%>
<script language=vbscript>
MsgBox "大侠，您目前的魅力低于10了，面子上看不过去呀，以后再来吧！"
location.href = "javascript:history.back()"
</script>
<%
rs.close
set rs=nothing
else
%>
<!--#include file="h_gamble.asp"-->
<%
Session("n_PokerBack")="deck_2.gif"
'Session("n_TotalMoney")=1000
Session("n_TotalMoney")=rs("银两")
Session("n_Bet")=0
Session("bet_win")=0
Session("n_Init")=false '第一次进入赌博页面标志
	
redim userPokers(1,0)
redim antiPokers(1,0)
		
Session("n_UserPokers")=userPokers '用户牌
Session("n_AntiPokers")=antiPokers '电脑牌
Session("n_Result")=4
rs.close
set rs=nothing
Response.Redirect "gamble.asp"
%>
<%
end if
%>
<html></html>