<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
cmid=Request.QueryString("id")
cmprice=Request.Form("price")
cmprice=clng(cmprice)
if cmprice<1 then cmprice=1
if not (isnumeric(cmid) and  isnumeric(cmprice))then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from ��Ʒ where id="&cmid&" and ������='"&username&"' and ����=True "
rst.open sqlstr,conn
if  rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=030"
rst.Close
set rst=nothing 	
sqlstr="update ��Ʒ set �۸�="&cmprice&" where id="&cmid&" and ������='"&username&"'"
conn.Execute sqlstr
conn.Close
set conn=nothing
%>
<head><title>���ɼ���</title><LINK href="../chatroom/css.css" rel=stylesheet><script language=javascript>setTimeout("location.replace('chgprice.asp');",3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
  <tr align=center> 
    <td><A href="market.asp" onmouseover="window.strtus='����������ͨ����Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
    <td><A href="sale.asp" onmouseover="window.strtus='��۳����Լ�ӵ�е���Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
    <td bgcolor=#FFCC00><A href="chgprice.asp" onmouseover="window.strtus='����Ѽ�����Ʒ�ļ۸�';return true;" onmouseout="window.status='';return true">�ļ�</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("Ba_jxqy_systemname")%>���ɼ���</b></font></p>
<div align="center">���ǻᰴ�µı�۳��۴���Ʒ������������ǽ����Ͻ�Ǯ���㣬ˡ������֪ͨ�ˣ��������10% </div>
<p align=center><a href="javascript:location.replace('chgprice.asp');" onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����</a></p>
</body>