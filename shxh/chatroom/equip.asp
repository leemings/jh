<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
commodityname=Request.QueryString("commodityname")
if instr(commodityname,"'")<>0 or instr(commodityname," ")<>0 then Response.End
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "chaterror.asp?id=000"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select tc.ID,tc.����,tc.����,tc.����,tc.��Ч,tu.�ؼ� from ��Ʒ tc inner join �û� tu on tu.����=tc.������ where tc.����='"&commodityname&"' and tc.������='"&username&"' and tc.����>0"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then
	msg="��û�д�����Ʒ�ɹ�װ��"
else
	cmid=rst("id")
	cmtype=rst("����")
	cmattack=rst("����")
	cmdefence=rst("����")
	cmespecial=rst("��Ч")
	especial=rst("�ؼ�")
	rst.Close
	rst.Open "select id from ��Ʒ where ����='"&cmtype&"' and װ��=true and ������='"&username&"'",conn
	if rst.EOF or rst.BOF then
		if cmespecial="��" or isnull(cmespecial) or cmespecial="" then
			especial=especial
		else 
			especial=especial&cmespecial&";"
		end if	
		conn.Execute "update �û� set ����=����+"&cmattack&",����=����+"&cmdefence&",�ؼ�='"&especial&"' where ����='"&username&"'"
		conn.Execute "update ��Ʒ set ����=����-1,װ��=True where id="&cmid
		msg="��ɹ���װ����"&commodityname&",������Щ������"&cmattack&",������˶�����"&cmdefence
	else
		msg="�������Ѿ�װ����"&cmtype&",�޷�����װ��"&commodityname
	end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing	
%>
<head><title>װ����Ʒ</title><link rel="stylesheet" href="style1.css"><script language=javascript>setTimeout("location.replace('seeequip.asp');",3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false >
<p align=center><font color=0000FF>װ����Ʒ</font></p>
<font color=FF0000><%=msg%></font>
<br><br><br>3���Ӻ��Զ�����<br><a href="javascript:location.replace('seeequip.asp');" onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����</a></body>