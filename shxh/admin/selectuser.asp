<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if not(session("Ba_jxqy_usercorp")="�ٸ�" and Session("Ba_jxqy_usergrade")=Application("Ba_jxqy_allright")) then Response.Redirect "../error.asp?id=500"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
%>
<head>
<link rel="stylesheet" href="../chatroom/css.css">
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=50 background="<%=bgimage%>" bgcolor="<%=bgcolor%>">
<div align=center>
��������ѯ<form action="showuser.asp" method=post name=form1><input type=text name=search maxlength=14 size=14 value=""><INPUT type=submit value=" �� �� " ></form>
��IP��ѯ<form action="showbyip.asp" method=post><input type=text name=ip maxlength=15 size=15 value=''><input type=submit value=' �� �� '></form>
���ȼ���ѯ<form action=showbygr.asp method=post>
<select name=grade>
	<option value=1 selected>1��</option>
	<option value=2>2��</option>
	<option value=3>3��</option>
	<option value=4>4��</option>
	<option value=5>5��</option>
	<option value=6>6��</option>
	<option value=7>7��</option>
	<option value=8>8��</option>
	<option value=9>9��</option>
	<option value=10>10��</option>
</select><input type=submit value='����'></form>
�����ɲ�ѯ<form action=showbyco.asp method=post><select name=corp>
<option value='��' selected>��������</option>
<%
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "����",conn
do while not rst.EOF
	Response.Write "<option value='"&rst("����")&"'>"&rst("����")&"</option>"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close 
set conn=nothing
%>
</select><input type=submit value='����'></form>
ɾ�������ʺ�<form action=deluser.asp method=post>
<select name=howday>
<option value=7>7��</option>
<option value=15>15��</option>
<option value=30 selected>30��</option>
<option value=90>90��</option>
</select>
<input type=submit value='ȷ��ɾ��'>
</form>
</div>
</body></html>
