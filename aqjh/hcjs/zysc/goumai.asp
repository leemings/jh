<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
cmid=request("id")
if not isnumeric(cmid) then
 cmid=1
end if
Set connb = Server.CreateObject("ADODB.Connection")
connb.Open application("aqjh_usermdb")
Set rst= Server.CreateObject("ADODB.Recordset")
sqlstr="select * from ���� where id="&cmid&" and ����>0"
rst.open sqlstr,connb
if rst.EOF or rst.BOF then
Response.Write "<script language=javascript>alert('�Բ���û���ҵ�������Ʒ!'); window.close();</script>"
Response.End
end if
syz=rst("������")
name=rst("����")
danjia=rst("�۸�")
sl=rst("����")
rst.Close
set rst=nothing
connb.Close
set connb=nothing
%>
<head><title>���ɼ���</title><LINK href="../../css.css" rel=stylesheet></head>
<body background="../../Bg.gif" oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp" onmouseover="window.strtus='����������ͨ����Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
<td bgcolor=#ffff00><A href="sale.asp" onmouseover="window.strtus='��۳����Լ�ӵ�е���Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("aqjh_chatroomname")%>���ɼ���</b></font></p>
<p align=center>������Ʒ������Ҫ�����ֽ𣬹���ɹ�������ֱ�ӻ���</p>
<div align=center>
<form action="gouok.asp?id=<%=cmid%>" method=post>
<table width='60%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center'>
<tr><td width=36>��Ʒ</td><td><%=name%></td></tr>
<tr><td>���</td><td><%=danjia%></td></tr>
<tr><td width=36>������</td><td><%=syz%></td></tr>
<tr><td>����</td><td><input type=text name=sl value='<%=sl%>' maxlength=4 size=9></td></tr>
<tr>
<td colspan=2 align=center height="32">
<input type=submit value=" �� �� "> <input type=reset value=" �� �� "> <input type=button onclick="javascript:history.back()" value=" �� �� "></td></tr>
</table>
</form>
</div></body>