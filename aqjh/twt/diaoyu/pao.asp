<%@ LANGUAGE=VBScript codepage ="936" %>
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
says="<font color=#ff0000>消息</font>"& aqjh_name &"在河边钓鱼：由于太不专心鱼跑了。。。"			'聊天数据
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
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title>鱼跑了</title></head>
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
很可惜，你的鱼溜走了......
</div>
</table><div align="center"><br>
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</div></td></tr></table></div></body></html>