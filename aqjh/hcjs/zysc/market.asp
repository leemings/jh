<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=cint(acpage)
if acpage<1 then acpage=1
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open application("aqjh_usermdb")
Set rst= Server.CreateObject("ADODB.Recordset")
sqlstr="select * from ���� where ����>0 order by �۸� "
rst.open sqlstr,conn,1,2
if  not (rst.EOF or rst.BOF) then
	rst.PageSize=10
	if acpage>rst.PageCount then acpage=rst.PageCount
	rst.AbsolutePage=acpage
	msg="<table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center'><tr bgcolor=ffff00 align=center><td  width=40>����</td><td width=100>����</td><td>�۸�</td><td>����</td><td>������</td><td width=40>����</td></TR>"
	for j=1 to rst.Pagesize
		if rst.EOF or rst.BOF then
			exit for
		else
		select case rst("����")
		case "w6"
			lx="��Ʒw6"
		case "w8"
			lx="��ҩw8"
		case "w9"
			lx="ħ��w9"
		case "w10"
			lx="����w10"
		case "w3"
			lx="װ��w3"
		end select
			msg=msg&"<tr><td>"&lx&"</td><td>"&rst("����")&"</td><td>"&rst("�۸�")&"</td><td>"&rst("����")&"</td><td>"&rst("������")&"</td><td align='center'><a href='goumai.asp?id="&rst("id")&"'>����</a></td></tr>"
			rst.MoveNext
		end if
	next
	msg=msg+"</tr></table><form action='market.asp' method=post id=form1 name=form1><table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center' bgcolor=ffFF00><tr align=center>"
	if acpage>1 then
		msg=msg&"<td><a href='market.asp?acpage=1' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='market.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='ǰһҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">ǰһҳ</a></td>"
	else
		msg=msg&"<td>��һҳ</td><td>ǰһҳ</td>"
	end if
	if acpage<rst.PageCount then
		msg=msg&"<td><a href='market.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='market.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='���һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">���һҳ</a></td>"
	else
		msg=msg&"<td>��һҳ</td><td>���һҳ</td>"
	end if
	msg=msg&"<td>��<input type=text name=acpage size=4 value='"&acpage&"'>ҳ��"&rst.PageCount&"ҳ<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"
end if
rst.Close
set rst=nothing 	
conn.Close 
set conn=nothing
if msg="" then msg="Ŀǰ������û�ж����ɹ����ۣ�"
%>
<head><title>���ɼ���</title><LINK href="../../css.css" rel=stylesheet><meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor='<%=bgcolor%>' background='../../bg.gif' oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
  <tr align=center> 
    <td bgcolor=#FFFF00><A href="market.asp" onmouseover="window.strtus='����������ͨ����Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
    <td><A href="sale.asp" onmouseover="window.strtus='��۳����Լ�ӵ�е���Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
    </tr></table>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("aqjh_chatroomname")%>���ɼ���</b></font></p>
<div align=center><%=msg%></div>
<p align=center>
  <input type="button" value=" �� �� " onClick="javascript:location.href='zysc.asp'" name="button"> 
  <input type="button" value=" �� �� " onClick="javascript:window.close();" name="button">
</p>
</body>