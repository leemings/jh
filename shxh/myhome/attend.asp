<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
pid=Request.QueryString("id")
if not isnumeric(pid) then Response.Redirect "../error.asp?id=024"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from pet where username='"&username&"' and option_T<"&nowtimetype&" and exist=true and id="&pid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=069"
msg="<tr><td><input type=hidden name=id value='"&pid&"'>����</td><td><input type=text name=biological value='"&Trim(rst("biological"))&"' size=14 maxlength=7></td></tr>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<html>
<head>
<title>����֮��</title>
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<form action=attend2.asp method=post>
<p align=center><font color=0000ff size=5>����֮��</font><br>�����ĵĺ���,���ǰ��Ľ���</p><hr>�����,����Ը���ĳ���ȡ���Լ�ϲ��������
<table width=50% align=center border=3>
<%=msg%>
<tr><td>����</td><td><select name=option>
	<option value='ιʳ' selected>ιʳ</option>
	<option value='ɢ��'>ɢ��</option>
	<option value='����'>����</option>
	<option value='˯��'>˯��</option>
	<option value='ϴ��'>ϴ��</option>
	</select></td></tr>
<tr><td>ʱ��</td><td><select name='howminute'>
	<option value='10'>10 �� ��</option>
	<option value='30'>�� С ʱ</option>
	<option value='60'>һ С ʱ</option>
	<option value='180'>�� С ʱ</option>
	<option value='720'>ʮ��Сʱ</option>
	<option value='1440'>һ&nbsp;&nbsp;&nbsp;&nbsp;��</option>
	</select>
</td></tr>
<tr><td colspan=2 align=center><input type=submit value='�չ�'> <input type=button value='����' onclick=history.back();></td></tr>
</table>
</form>
</body>
</html>