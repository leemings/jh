<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
conn.execute "update 用户 set 操作时间=now() where 姓名='"&aqjh_name&"'"
conn.close
set conn=nothing
Session("diaoyu")=true
ff="<font color=#ff0000>消息</font>"& aqjh_name &"在后山挖宝藏：由于强盗人多势众,抢走了到手的财宝."			'聊天数据
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title>被抢了</title></head>
<body background="../../bg.gif">
<div align="center">
<p>&nbsp;</p>
<table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="200">
<tr>
<td bgcolor=#C6BD9B>
<table>
<tr>
<td valign="top">
<div align="center">
<font color="#000000"><img border="0" src="../../chat/img/Error.gif">
好痛苦,你快要到手的宝藏被强盗抢走了...</font>
</div>
</table>
<div align="center"><br>
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</div></td></tr></table></div></body></html>