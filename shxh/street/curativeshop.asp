<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=cint(acpage)
if acpage<1 then acpage=1
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select ID,����,����,����,��Ч,�۸� from �̵� where ����='ҩƷ' order by �۸� "
rst.open sqlstr,conn,1,2
if  not (rst.EOF or rst.BOF) then
	rst.PageSize=10
	if acpage>rst.PageCount then acpage=rst.PageCount
	rst.AbsolutePage=acpage
	msg="<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF><tr bgcolor=ffFF00 align=center><td width=100>����</td><td>����</td><td>����</td><td>��Ч</td><td>�۸�</td></TR>"
	for j=1 to rst.Pagesize
		if rst.EOF or rst.BOF then
			exit for
		else	
			msg=msg&"<tr><td><a href='buy.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='����"&rst("����")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("��Ч")&"</td><td>"&rst("�۸�")&"</td></tr>"
			rst.MoveNext
		end if
	next
	msg=msg+"</table><form action='curativeshop.asp' method=post><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=ffff00><tr align=center>"
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
<head><LINK href="../chatroom/css.css" rel=stylesheet><title>ҩ��</title></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("Ba_jxqy_systemname")%>ҩƷ��</b></font></p>
<div align=center><%=msg%></div>
<p align=center>
  <input type="button" value=" �� �� " onClick="javascript:location.href='street.asp'" name="button"> 
  <input type="button" value=" �� �� " onClick="javascript:window.close();" name="button">
</p>

</body>