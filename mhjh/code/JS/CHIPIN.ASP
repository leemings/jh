<%Response.Expires=-1
username=session("yx8_mhjh_username")
if username=""  then  Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from �û� where ����='"&username&"'"
rst.open sqlstr,conn
umoney=rst("����")
rst.Close
set rst=nothing
conn.close
set conn=nothing
%>
<head><title></title><LINK rel=stylesheet href='../style4.css'></head>
<body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'>
<form method="post" action="compete.asp"  target=compfrm>
<table  border="0" cellspacing="0" cellpadding="0" width='95%' align=center>
<tr>
<td width="27">0��</td>
<td width="34">
<input  maxlength=6 size=4  name="horse0" >
</td>
<td width="24">1�� </td>
<td width="34">
<input  maxlength=6 size=4  name="horse1" >
</td>
<td width="24">2�� </td>
<td width="34">
<input  maxlength=6 size=4  name="horse2" >
</td>
<td width="24">3�� </td>
<td width="34">
<input  maxlength=6 size=4  name="horse3" >
</td>
<td width="24">4��</td>
<td width="34">
<input  maxlength=6 size=4  name="horse4" >
</td>
<td width="24">5�� </td>
<td width="34">
<input  maxlength=6 size=4  name="horse5" >
</td>
<td width="24">6�� </td>
<td width="34">
<input  maxlength=6 size=4  name="horse6" >
</td>
<td width="24">7��</td>
<td width="34">
<input  maxlength=6 size=4  name="horse7" >
</td>
<td width="24">8�� </td>
<td width="34">
<input  maxlength=6 size=4  name="horse8" >
</td>
<td width="24">9��</td>
<td width="34">
<input  maxlength=6 size=4  name="horse9" >
</td>
<td width="24">10��</td>
<td width="34">
<input  maxlength=6 size=4  name="horse10" >
</td>
<td width="24">11�� </td>
<td width="34">
<input  maxlength=6 size=4  name="horse11" >
</td>
</tr>
<tr><td colspan=12 align=right><input type=submit value=" �� ע "> <input type=button onclick=javascript:top.window.close(); value=' �� �� '> </td><td colspan=12><%Response.Write username&"��������<input type=text readonly value='"&umoney&"'>"%></td></tr>
</table>
</FORM>
</body>
