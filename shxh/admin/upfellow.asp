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
if not(usercorp="�ٸ�" and usergrade=Application("Ba_jxqy_allright")) then Response.Redirect "../error.asp?id=500"
st=server.htmlencode(trim(Request.Form("username")))
mn=server.htmlencode(trim(Request.Form("month")))
opt=server.htmlencode(trim(Request.Form("submit")))
if instr(st,"'")<>0 then Response.Redirect "../error.asp?id=024"
if not isnumeric(mn) then Response.Redirect "../error.asp?id=024"
today=date()
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ��Աʱ�� from �û� where ����='"&st&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=228"
if rst("��Աʱ��")>today then today=rst("��Աʱ��")
if opt="�� ��" then
	fellowdate=dateadd("m",mn,today)
else
	fellowdate=dateadd("m",-mn,today)
end if	
fellowdatetype="#"&month(fellowdate)&"/"&day(fellowdate)&"/"&year(fellowdate)&"#"
rst.Close
set rst=nothing
conn.Execute "update �û� set ��Ա=true,��Աʱ��="&fellowdatetype&" where ����='"&st&"'"
conn.Close
set conn=nothing
%>
<html>
<head>
<link rel=stylesheet href='../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<p align=center><font color=0000ff size=4>��Ա����</font></p><hr>
<table width=80% align=center>
<tr><td>��Ա</td><td><%=st%></td></tr>
<tr><td>ʱ��</td><td><%=formatdatetime(fellowdate,1)%></td></tr>
<tr align=center><td colspan=2><a href='showfell.asp' onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����</a></td></tr>
</table>
</body>
</html>