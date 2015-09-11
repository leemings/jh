
<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp" -->
<!--#include file="../../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
'call bdhy()
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示:你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
act=1
nowinroom=session("nowinroom")
addwordcolor="660099"
saycolor="008888"
towho="大家"
addsays=aqjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs1=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
'rs.open "SELECT * from aqjh_system where id=502",conn,1,1
'diaoyusj=clng(rs("svalue"))
'rs.close
rs.open "select * from 用户 where 姓名='" & aqjh_name & "'",conn,1,3
if rs("times")>3 then
rs.close
set rs=nothing
Response.Write "<script language=JavaScript>alert('提示：您已经失去领取资格了,领取条件只限第一次登陆的用户！');window.close();</script>"
Response.End
end if
	Response.Write "<script Language=Javascript>function sf(){};</script>"
	mess2="<font color=red>"&aqjh_name&"</font>得到站长准备的[福神卡]<font color=#ff0000>1张</font>、[涨钱卡]<font color=#ff0000>1张</font> [<font color=red>领取条件为第一次登陆本江湖,而且只能领取一次!</font>]"
	rs1.open "SELECT w5 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	kapian=add(rs1("w5"),"福神卡",1)
	conn.execute "update 用户 set  w5='"&kapian&"' where 姓名='"&aqjh_name&"'"
	rs2.open "SELECT w5 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	kapian2=add(rs2("w5"),"涨钱卡",1)
	conn.execute "update 用户 set  w5='"&kapian2&"' where 姓名='"&aqjh_name&"'"
    'conn.Execute "update 用户 set 银两=银两+5000000,金=金+10,木=木+10,水=水+10,火=火+10,土=土+10,times=times+1 where 姓名='"&aqjh_name&"'"
	conn.Execute "update 用户 set times=times+1 where 姓名='"&aqjh_name&"'"
	'rs("jh_czsj")=now()
	rs.update
	says="<img src=img/xinzhou.gif><font color=#ff0000>【新人费】</font>"&aqjh_name&"领取新人费：" & mess2
    call showchat(says)
	'show="<table border=1 bgcolor=#948754 align=center cellpadding=10 cellspacing=13 width=100% height=100% ><tr><td align=center bgcolor=#C6BD9B>"& replace(mess2,"##",aqjh_name) &"<br><br><br><input type=button value=关闭窗口 onClick='window.close()' name=button></td></tr></table>"
conn.close
set conn=nothing
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title>领取新人费OK</title>
<link href="../../css.css" rel="stylesheet" type="text/css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#000000">
<div align="center"> <br>
<br>
<table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="186" width="300">
<tr>
<td bgcolor=#C6BD9B>
<table align="center" width="248">
<tr>
<td valign="top">
<div align="center">
<p><%=mess2%></p>
</div>
</table>
<div align="center"><br>
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</div>
</td>
</tr>
</table>
</div>
</body>
</html>