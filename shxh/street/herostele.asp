<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=cint(acpage)
if acpage<1 then acpage=1
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from Ӣ���� order by ʱ�� desc"
rst.open sqlstr,conn,1,2
if  not (rst.EOF or rst.BOF) then
	rst.PageSize=20
	if acpage>rst.PageCount then acpage=rst.PageCount
	rst.AbsolutePage=acpage
	msg="<table border=3 width='100%' bgcolor=666666 cellpadding='3' cellspacing='1' align=center><tr bgcolor=00FF00 align=center><td>ʱ��</td><td>����</td><td>����</td><td>����</td></TR>"
	for j=1 to rst.Pagesize
		if rst.EOF or rst.BOF then
			exit for
		else	
			msg=msg&"<tr bgcolor=cccc99><td>"&rst("ʱ��")&"</td><td>"&rst("����")&"</td><td align=right>"&rst("����")&"</td><td>"&rst("����")&"</td></tr>"
			rst.MoveNext
		end if
	next
	msg=msg+"</tr></table><form action='herostele.asp' method=post id=form1 name=form1><table border=1 width=100% bgcolor=666666 cellpadding='3' cellspacing='1' align=center><tr align=center bgcolor=ffff00>"
	if acpage>1 then
		msg=msg&"<td><a href='herostele.asp?acpage=1' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='herostele.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='ǰһҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">ǰһҳ</a></td>"
	else
		msg=msg&"<td>��һҳ</td><td>ǰһҳ</td>"
	end if
	if acpage<rst.PageCount then
		msg=msg&"<td><a href='herostele.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='herostele.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='���һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">���һҳ</a></td>"
	else
		msg=msg&"<td>��һҳ</td><td>���һҳ</td>"
	end if
	msg=msg&"<td>��<input type=text name=acpage size=4 value='"&acpage&"'>ҳ/��"&rst.PageCount&"ҳ<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"
end if
rst.Close
set rst=nothing 	
conn.Close 
set conn=nothing
if msg="" then msg="Ŀǰû����������"
%>
<head><title>Application("Ba_jxqy_systemname")</title><link rel="stylesheet" href="../chatroom/css.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("Ba_jxqy_systemname")%>Ӣ����</b></font></p>
<div align=center><%=msg%></div>
<p align=center><a href="#" onclick=javascript:window.close(); onmouseover="window.status='�ر�';return true;" onmouseout="window.status='';return true;">�ر�</a></p>
</body>