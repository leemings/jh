<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
cmid=Request.QueryString("id")
if not isnumeric(cmid) then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.Mode=16
conn.IsolationLevel=256
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from ��Ʒ where id="&cmid&" and ����=True "
rst.open sqlstr,conn
if  rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=029"
cmname=rst("����")
cmtype=rst("����")
cmhp=rst("����")
cmmp=rst("����")
cmattack=rst("����")
cmdefence=rst("����")
cmespecial=rst("��Ч")
cmprice=rst("�۸�")
cmhost=rst("������")
rst.Close
rst.Open "select * from �û� where ����='"&username&"' and ����>="&cmprice,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=030"
rst.Close
conn.BeginTrans
conn.Execute "update ��Ʒ set ����=False where id="&cmid
conn.CommitTrans
rst.Open "select * from ��Ʒ where ����='"&cmname&"' and ������='"&username&"'",conn
if rst.EOF or rst.BOF then
	conn.Execute "insert into ��Ʒ(����,����,����,����,����,����,��Ч,�۸�,����,������,����,װ��) values('"&cmname&"','"&cmtype&"',"&cmhp&","&cmmp&","&cmattack&","&cmdefence&",'"&cmespecial&"',"&cmprice&",1,'"&username&"',False,False)"
else
	conn.Execute "update ��Ʒ set ����=����+1 where ����='"&cmname&"' and ������='"&username&"'"
end if
conn.Execute "update �û� set ����=����+"&cmprice*0.9&" where ����='"&cmhost&"'"
conn.Execute "update �û� set ����=����-"&cmprice&" where ����='"&username&"'"
%>
<head><title>���ɼ���</title><LINK href="../chatroom/css.css" rel=stylesheet><script language=javascript>setTimeout("location.replace('market.asp');",3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
  <tr align=center> 
    <td bgcolor=#FFFF00><A href="market.asp" onmouseover="window.strtus='����������ͨ����Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
    <td><A href="sale.asp" onmouseover="window.strtus='��۳����Լ�ӵ�е���Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
    <td><A href="chgprice.asp" onmouseover="window.strtus='����Ѽ�����Ʒ�ļ۸�';return true;" onmouseout="window.status='';return true">�ļ�</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("Ba_jxqy_systemname")%>���ɼ���</b></font></p>
<div align="center">Ǯ�����������պ��ˣ�лл���Ļݹ� </div>
<p align=right><a href="javascript:location.replace('market.asp');" onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����</a></p>
</body>