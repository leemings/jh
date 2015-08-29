<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
'if InStr(http,"dalie/fl.asp")=0 then 
'Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
'Response.End 
'end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	response.end
end if
if Application("sjjh_dalie")<>"老虎" then
	Response.Write "<script Language=Javascript>alert('提示：现在还没有猎物可以打，请一会再来吧。');window.close();</script>"
	response.end
end if
Application.Lock
Application("sjjh_dalie")="无"
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT w6 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3
duyao=add(rs("w6"),"虎肉",1)
conn.execute "update 用户 set  w6='"&duyao&"',内力=内力+800,体力=体力+1200,allvalue=allvalue+1,银两=银两+5000 where 姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>【好消息】</font><b><font color=red>野兽已经被'"&sjjh_name&"'打死，真是英雄少年呀，请大家放心聊天吧。</font></b>"		'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
<html>
<title>打猎成功</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
-->
</style>

<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<table width="98%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
<tr>
<td height="17" bgcolor="#0000FF" align="center"><font color="#FFFFFF">打猎成功</font></td>
</tr>
<tr>
<td bgcolor="#FFCC99" align="center" height="157"><font> 您成功的将野兽打死，增加内力+800，体力+1200，银两+5000，点券+1。<br>
<br>
虎肉一块，使用物品:内力+150,体力+150</font></td>
</tr>
</table>
<p>&nbsp;</p>
</body>
</html>
