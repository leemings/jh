<%@ LANGUAGE=VBScript codepage ="936" %><%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
lx=request("lx")
name=request("name")
sl=request("sl")
danjia=request("danjia")
select case lx
case "��Ʒ"
dang="w6"
case "��ҩ"
dang="w8"
case "ħ��"
dang="w9"
case "����"
dang="w10"
case "װ��"
dang="w3"
case else
Response.Write "<script language=javascript>alert('�Բ������ͳ���,��������?'); window.close();</script>"
Response.End
end select
if name="" then
Response.Write "<script language=javascript>alert('�Բ�����Ʒ������!'); window.close();</script>"
Response.End
end if
if not isnumeric(sl) or not isnumeric(danjia) then
Response.Write "<script language=javascript>alert('�Բ��������򵥼۳���!'); window.close();</script>"
Response.End
end if
%>
<head><title>���ɼ���</title><LINK href="../../css.css" rel=stylesheet></head>
<body background="../../bg.gif" oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp">����</a></td>
<td bgcolor=#ffff00><A href="sale.asp">����</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("aqjh_chatroomname")%>���ɼ���</b></font></p>
<p align=center>ע���̵�����򵽵Ķ�����ֹ����</p>
<div align=center>
<form action="sell.asp" method=post>
<input name="lx" type="hidden" value='<%=lx%>'>
<input name="name" type="hidden" value='<%=name%>'>
<table width='60%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center'>
<tr><td width=36>����</td><td><%=lx%>��</td></tr>
<tr><td width=36>��Ʒ</td><td><%=name%></td></tr>
<tr><td>����</td><td><input type=text name=danjia value='<%=danjia%>' maxlength=9 size=9></td></tr>
<tr><td>����</td><td><input type=text name=sl value='<%=sl%>' maxlength=4 size=9></td></tr>
<tr>
<td colspan=2 align=center height="32">
<input type=submit value=" �� �� "> <input type=reset value=" �� �� "> <input type=button onclick="javascript:history.back()" value=" �� �� "></td></tr>
</table>
</form>
</div></body>