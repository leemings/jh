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
sqlstr="select * from ��Ʒ where ������='"&username&"' and ����=True order by �۸� "
rst.open sqlstr,conn,1,2
if  not (rst.EOF or rst.BOF) then
	rst.PageSize=10
	if acpage>rst.PageCount then acpage=rst.PageCount
	rst.AbsolutePage=acpage
	msg="<table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center'><tr bgcolor=ffFF00 align=center><td width=40>����</td><td width=140>����</td><td>�ο��۸�</td><td width=36>����</td></TR>"
	for j=1 to rst.Pagesize
		if rst.EOF or rst.BOF then
			exit for
		else	
			msg=msg&"<tr><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td align=right>"&rst("�۸�")&"</td><td><a href='changeprice.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='����"&rst("����")&"�ļ۸�';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">����</a></td></tr>"
			rst.MoveNext
		end if
	next
	msg=msg+"</tr></table><form action='chgprice.asp' method=post id=form1 name=form1><table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center' bgcolor=ffFF00><tr align=center>"
	if acpage>1 then
		msg=msg&"<td><a href='chgprice.asp?acpage=1' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='chgprice.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='ǰһҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">ǰһҳ</a></td>"
	else
		msg=msg&"<td>��һҳ</td><td>ǰһҳ</td>"
	end if
	if acpage<rst.PageCount then
		msg=msg&"<td><a href='chgprice.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='chgprice.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='���һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">���һҳ</a></td>"
	else
		msg=msg&"<td>��һҳ</td><td>���һҳ</td>"
	end if
	msg=msg&"<td>��<input type=text name=acpage size=4 value='"&acpage&"'>ҳ��"&rst.PageCount&"ҳ<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"
end if
rst.Close
set rst=nothing 	
conn.Close 
set conn=nothing
if msg="" then msg="��û�ж����ɹ����ģ�"
%>
<head><title>���ɼ���</title><LINK href="../chatroom/css.css" rel=stylesheet></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
  <tr align=center> 
    <td><A href="market.asp" onmouseover="window.strtus='����������ͨ����Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
    <td><A href="sale.asp" onmouseover="window.strtus='��۳����Լ�ӵ�е���Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
    <td bgcolor=#FFFF00><A href="chgprice.asp" onmouseover="window.strtus='����Ѽ�����Ʒ�ļ۸�';return true;" onmouseout="window.status='';return true">�ļ�</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("Ba_jxqy_systemname")%>���ɼ���</b></font></p>
<div align=center><%=msg%></div>
<p align=center>
  <input type="button" value=" �� �� " onClick="javascript:location.href='street.asp'" name="button"> 
  <input type="button" value=" �� �� " onClick="javascript:window.close();" name="button">
</p>
</body>