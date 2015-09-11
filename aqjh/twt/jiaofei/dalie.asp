<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
'if InStr(http,"twt/dalie/fl.asp")=0 then 
'Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
'Response.End 
'end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	response.end
end if
if Application("aqjh_dalie")<>"土匪" then
	Response.Write "<script Language=Javascript>alert('提示：现在还没有土匪。');window.close();</script>"
	response.end
end if
Application.Lock
Application("aqjh_dalie")="无"
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT w6 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
duyao=add(rs("w6"),"矿石",3)
conn.execute "update 用户 set  w6='"&duyao&"',内力=内力+1000,体力=体力+6000,allvalue=allvalue+1,银两=银两+5000 where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>【好消息】</font><font color=red>土匪已经被'"&aqjh_name&"'奋力搏杀，真是英雄少年呀，打扫战场后发现三块矿石。。</font>"		'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function

Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>
<html><head>
<title>剿匪成功</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
-->
</style>
</head>
<body text="#000000" background="../../bg.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF" ">
<table width="98%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
<tr>
<td height="17" bgcolor="#0000FF" align="center"><font color="#FFFFFF">剿匪成功</font></td>
</tr>
<tr>
<td bgcolor="#FFCC99" align="center" height="157"><font> 您成功的将土匪打死，增加内力+8000，体力+6000，银两+5000，战斗经验+2。<br>
<br>
打扫战场后发现三块矿石，</font></td>
</tr></table><p>&nbsp;</p></body></html>