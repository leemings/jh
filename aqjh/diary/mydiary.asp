<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if request("name")="" then
name=aqjh_name
else
name=request("name")
end if
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
rs1.Open "Select * From 日记用户 where 姓名='"& name &"'", conn1, 1,1
if rs1.eof and rs1.bof then
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing
Response.Redirect "reg.asp"
end if
rjname=rs1("日记本名")
sm=rs1("说明")
lb1=rs1("lb1")
lb2=rs1("lb2")
lb3=rs1("lb3")
rjs=rs1("日记数")
jl=rs1("建立日期")
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing
%>
<html><head><title><%=name%>的日记本</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head><LINK href="images/diary.css" rel=stylesheet type=text/css>
<body bgcolor="#ebe4de" text="#FFFFFF"><table align=center border=0 cellpadding=0 cellspacing=0 width=557><tr>
<td align=center background=images/bg.gif height=352 valign=top>
<table border=0 cellpadding=3 cellspacing=0 width="50%"><tr><td height=110 valign="bottom" align="left" colspan="2"><font color=#ffffff>日记本名：<%=rjname%></font></td>
</tr><tr><td height=80 valign="top" align="left" colspan="2"> <font color=#ffffff>首页寄语：<br><%=sm%></font><br>
<br><br></td></tr><tr><td colspan="2"><font color=#ffffff>日记拥有者：<%=name%>&nbsp;日记数：<%=rjs%></font></td>
</tr><tr><td colspan="2"><font color=#ffffff>建立时间：<%=jl%></font></td></tr><tr>
<td align=left valign="top" width="66%">&nbsp;</td><td align=right width="34%"><a href="main.asp?name=<%=name%>"><img alt=进入日记本 border=0 height=37 src="images/jinru.gif" width=94></a>
</td></tr></table></td></tr></table></body></html>