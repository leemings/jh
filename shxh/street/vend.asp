<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
cmid=Request.QueryString("id")
if not isnumeric(cmid) then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select ID,����,�۸� from ��Ʒ where id="&cmid&" and ������='"&username&"' and ����>0 and ����=False "
rst.open sqlstr,conn
if  rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=028"
cmname=rst("����")
cmprice=rst("�۸�")
rst.Close
set rst=nothing 	
conn.Close 
set conn=nothing
%>
<head><title>���ɼ���</title><LINK href="../chatroom/css.css" rel=stylesheet></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
  <tr align=center> 
    <td><A href="market.asp" onmouseover="window.strtus='����������ͨ����Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
    <td bgcolor=#FFFF00><A href="sale.asp" onmouseover="window.strtus='��۳����Լ�ӵ�е���Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
    <td><A href="chgprice.asp" onmouseover="window.strtus='����Ѽ�����Ʒ�ļ۸�';return true;" onmouseout="window.status='';return true">�ļ�</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("Ba_jxqy_systemname")%>���ɼ���</b></font></p>
<div align=center>
<form action="sell.asp?id=<%=cmid%>" method=post>
<table width='60%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center'>
	<tr><td width=36>��Ʒ</td><td><%=cmname%></td></tr>
	<tr><td>���</td><td><input type=text name=price value='<%=cmprice%>' maxlength=9 size=9></td></tr>
	<tr>
        <td colspan=2 align=center height="32">
<input type=submit value=" �� �� "> <input type=reset value=" �� �� "> <input type=button onclick="javascript:history.back()" value=" �� �� "></td></tr>
</table>
</form>
</div>
<p align=center>
  <input type="button" value=" �� �� " onClick="javascript:top.window.close();" name="button">
</p>
</body>