<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if not(session("Ba_jxqy_usercorp")="�ٸ�" and Session("Ba_jxqy_usergrade")=Application("Ba_jxqy_allright")) then Response.Redirect "../error.asp?id=500"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("Ba_jxqy_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
rst.Open "ϵͳ����",conn
if rst.EOF or rst.BOF then	Response.Redirect "../error.asp?id=100"
%>
<head>
<link rel="stylesheet" href="../style.css">
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0 background="<%=bgimage%>" bgcolor="<%=bgcolor%>">
<form method=post action=updatesystem.asp id=form1 name=form1>
	<table width=100% border=3>
		<tr><td width=30%>ϵͳ����</td><td>����ֵ</td></tr>
<%
do while not rst.EOF%>
	<tr><td><%=rst("����")%></td><td><input type=text name="<%=rst("����")%>" value="<%=rst("����ֵ")%>"  size=50 maxlength=300></td></tr>
<%	
rst.MoveNext
loop%>
		<tr><td colspan=2=2 align=center><input type=submit value=' �� ��' > <input type=reset value=' �� �� '></td></tr>
	</table>
</form>
</body></html>
<%
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>