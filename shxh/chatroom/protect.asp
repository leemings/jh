<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
onlinenum=Application("Ba_jxqy_allonlinenum")
if Application("Ba_jxqy_fellow")=true then msg="��Ա���ܿ�����Ŀǰֻ�л�Ա�����뱣����"
username=session("Ba_jxqy_username")
chatroomsn=session("Ba_jxqy_userchatroomsn")
chatroomname=Application("Ba_jxqy_systemname"&chatroomsn)
if username="" then Response.Redirect "../error.asp?id=016"
%>
<html>
<head>
<link rel=stylesheet href='css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=ngimage%>'>
<p align=center>
<font size=5 color=0000ff><%=chatroomname%><br>
<font color=ff0000><%=onlinenum%></font>������
<hr>
���뱣��</font></p>
<%=msg%>
�����ڼ䣬�㽫�޷��������ˣ�Ҳ���������˹�����Ҳ�޷�����ܱ���״̬��
<form action= pronow.asp method=post>
<table align=center>
<tr>
	<td>
		<select name='howminute'>
		<option value='10'>10 �� ��</option>
		<option value='30'>�� С ʱ</option>
		<option value='60'>һ С ʱ</option>
		<option value='180'>�� С ʱ</option>
		<option value='720'>ʮ��Сʱ</option>
		<option value='1440'>һ&nbsp;&nbsp;&nbsp;&nbsp;��</option>
		</select>
	</td>
</tr>
<tr>
	<td>
		<input type=submit value='����'>
	</td>
</tr>
</table>
</form>
</html>