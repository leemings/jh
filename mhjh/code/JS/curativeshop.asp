<%
Response.Expires=-1
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=cint(acpage)
if acpage<1 then acpage=1
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select ID,����,����,����,��Ч,�۸� from �̵� where ����='ҩƷ' order by �۸� "
rst.open sqlstr,conn,1,2
if  not (rst.EOF or rst.BOF) then
rst.PageSize=10
if acpage>rst.PageCount then acpage=rst.PageCount
rst.AbsolutePage=acpage
msg="<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=ffffff><tr bgcolor=ffffff align=center><td width=100>����</td><td>����</td><td>����</td><td>��Ч</td><td>�۸�</td><td>����</td><td>����</td></TR>"
for j=1 to rst.Pagesize
if rst.EOF or rst.BOF then
exit for
else
msg=msg&"<tr><form method=POST action='buy.asp'><td><input type=hidden checked value='"&rst("id")&"' name='id'>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("��Ч")&"</td><td>"&rst("�۸�")&"</td><td><select name='sl'><option value='1' selected>1</option><option value='2'>2</option><option value='3'>3</option><option value='4'>4</option><option value='5'>5</option><option value='6'>6</option><option value='7'>7</option><option value='8'>8</option><option value='9'>9</option></select></td><td><input type='SUBMIT' name='Submit' value='����'></td></form></tr>"
rst.MoveNext
end if
next
msg=msg+"</table><form action='curativeshop.asp' method=post><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=ffffff bgcolor=ffffff><tr align=center>"
if acpage>1 then
msg=msg&"<td><a href='curativeshop.asp?acpage=1' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='curativeshop.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='ǰһҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">ǰһҳ</a></td>"
else
msg=msg&"<td>��һҳ</td><td>ǰһҳ</td>"
end if
if acpage<rst.PageCount then
msg=msg&"<td><a href='curativeshop.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='curativeshop.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='���һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">���һҳ</a></td>"
else
msg=msg&"<td>��һҳ</td><td>���һҳ</td>"
end if
msg=msg&"<td>��<input type=text name=acpage size=4 value='"&acpage&"'>ҳ��"&rst.PageCount&"ҳ<input type=submit value='GO'></td></tr></table></form>"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><LINK href="../style.css" rel=stylesheet><title>ҩ��</title></head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false>
<p align=center><b><font face="��������" size="3">ɽ��ҩ��</font></b></p>
<div align=center><%=msg%></div>
<p align=center>
<input type="button" value=" �� �� " onClick="javascript:window.close();" name="button">
</p>


</body>

