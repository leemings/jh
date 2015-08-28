<%
Response.Buffer =true
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"

st=server.htmlencode(trim(Request.Form("username")))
mn=server.htmlencode(trim(Request.Form("month")))
opt=server.htmlencode(trim(Request.Form("submit")))
if instr(st,"'")<>0 then Response.Redirect "../error.asp?id=024"
if not isnumeric(mn) then Response.Redirect "../error.asp?id=024"
today=date()
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 会员时间 from 用户 where 姓名='"&st&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=228"
if rst("会员时间")>today then today=rst("会员时间")
if opt="增 加" then
fellowdate=dateadd("m",mn,today)
else
fellowdate=dateadd("m",-mn,today)
end if
fellowdatetype="#"&month(fellowdate)&"/"&day(fellowdate)&"/"&year(fellowdate)&"#"
rst.Close
set rst=nothing
conn.Execute "update 用户 set 会员=true,会员时间="&fellowdatetype&" where 姓名='"&st&"'"
conn.Close
set conn=nothing
%>
<html>
<head>
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'>
<p align=center><font color=0000ff size=4>会员管理</font></p><hr>
<table width=80% align=center>
<tr><td>会员</td><td><%=st%></td></tr>
<tr><td>时限</td><td><%=formatdatetime(fellowdate,1)%></td></tr>
<tr align=center><td colspan=2><a href='showfell.asp' onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></td></tr>
</table>
</body>
</html>
