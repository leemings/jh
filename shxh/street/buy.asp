<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
commodityid=Request.QueryString("id")
if isnumeric(commodityid) then
	set conn=server.CreateObject("adodb.connection")
	conn.Open Application("Ba_jxqy_connstr")
	set rst=server.CreateObject("adodb.recordset")
	sqlstr="select * from �̵� where id="&commodityid
	rst.Open sqlstr,conn
	if rst.EOF or rst.BOF then
		msg="���󣬲����۴���Ʒ��"
	else
		commodityname=rst("����")
		commoditytype=rst("����")
		commodityhealth=rst("����")
		magic=rst("����")
		attack=rst("����")
		defence=rst("����")
		especial=rst("��Ч")
		commodityprice=rst("�۸�")
		rst.Close
		rst.Open "select ���� from �û� where ����='"&username&"' and ����>="&commodityprice,conn
		if rst.EOF or rst.BOF then
			msg="�������Ǯ����û�д���Ӵ��"
		else
			rst.Close
			sqlstr="select * from ��Ʒ where ����='"&commodityname&"' and ������='"&username&"'"
			rst.Open sqlstr,conn
			if rst.EOF or rst.BOF then
				sqlstr="insert into ��Ʒ(����,����,����,����,����,����,��Ч,�۸�,����,������,����,װ��) values('"&commodityname&"','"&commoditytype&"',"&commodityhealth&","&magic&","&attack&","&defence&",'"&especial&"',"&commodityprice/2&",1,'"&username&"',False,False)"
			else
				sqlstr="update ��Ʒ set ����=����+1 where ����='"&commodityname&"' and ������='"&username&"'"
			end if
			conn.BeginTrans
			conn.Execute sqlstr
			conn.Execute "update �û� set ����=����-"&commodityprice&" where ����='"&username&"'"
			conn.CommitTrans
			msg="�л�����"&commodityprice&"��������"&commodityname&"�����պ��ˣ�"
		end if	
	end if
	rst.Close
	set rst=nothing
	conn.Close
	set conn=nothing	
else
	msg="���󣬲����۴���Ʒ��"
end if
%>
<head><title>�̵�</title><LINK href="../chatroom/css.css" rel=stylesheet><script language=javascript>setTimeout('history.back();',3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false topMargin=200>
<div align=center>3���Ӻ��Զ�����<br>
  <font color=FF0000><img src="../images/error.gif" width="32" height="32"> <%=msg%></font></div>
<p align=right><a href="#" onclick=javascript:history.back(); onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����</a></p>
</body>