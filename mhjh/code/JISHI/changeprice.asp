<%
Response.Expires=-1
bgcolor=Application("yx8_mhjh_backgroundcolor")
cmid=Request.QueryString("id")
if not isnumeric(cmid) then Response.Redirect "../error.asp?id=024"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select ID,����,�۸� from ��Ʒ where id="&cmid&" and ������='"&username&"' and ����=True "
rst.open sqlstr,conn
if  rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=030"
cmname=rst("����")
cmprice=rst("�۸�")
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><title>���ɼ���</title><LINK href="../style4.css" rel=stylesheet></head>
<body bgcolor='<%=bgcolor%>'  oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp" onmouseover="window.strtus='����������ͨ����Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
<td><A href="sale.asp" onmouseover="window.strtus='��۳����Լ�ӵ�е���Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
<td bgcolor=#005b00><A href="chgprice.asp" onmouseover="window.strtus='����Ѽ�����Ʒ�ļ۸�';return true;" onmouseout="window.status='';return true">�ļ�</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("yx8_mhjh_systemname")%>���ɼ���</b></font></p>
<div align=center>
<form action="changenow.asp?id=<%=cmid%>" method=post >
<table border=1 width=50% cellpadding="3" cellspacing="1" bordercolor="#FF9900">
<tr><td width=36>��Ʒ</td><td><%=cmname%></td></tr>
<tr><td>���</td><td><input type=text name=price value='<%=cmprice%>' maxlength=9 size=9></td></tr>
<tr><td colspan=2 align=center><input type=submit value=" �� �� " id=submit1 name=submit1> <input type=reset value=" �� �� " id=reset1 name=reset1> <input type=button onclick="javascript:history.back()" value=" �� �� " id=button1 name=button1></td></tr>
</table>
</form>
</div>
<p align=center><input type="button" value=" �� �� " onClick="javascript:top.window.close();" name="button"></p>
</body>