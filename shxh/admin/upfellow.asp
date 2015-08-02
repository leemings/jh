<%
Response.Buffer =true
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("Ba_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usercorp="官府" and usergrade=Application("Ba_jxqy_allright")) then Response.Redirect "../error.asp?id=500"
st=server.htmlencode(trim(Request.Form("username")))
mn=server.htmlencode(trim(Request.Form("month")))
opt=server.htmlencode(trim(Request.Form("submit")))
if instr(st,"'")<>0 then Response.Redirect "../error.asp?id=024"
if not isnumeric(mn) then Response.Redirect "../error.asp?id=024"
today=date()
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
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
<link rel=stylesheet href='../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<p align=center><font color=0000ff size=4>会员管理</font></p><hr>
<table width=80% align=center>
<tr><td>会员</td><td><%=st%></td></tr>
<tr><td>时限</td><td><%=formatdatetime(fellowdate,1)%></td></tr>
<tr align=center><td colspan=2><a href='showfell.asp' onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></td></tr>
</table>
</body>
</html>