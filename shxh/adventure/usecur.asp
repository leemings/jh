<%
username=session("Ba_jxqy_username")
cur=Request.QueryString("cur")
if instr(cur,"'")<>0 then Response.End
if username="" then 
	msg="top.location.location.href='error.asp?id=016';"
else
	set conn=server.CreateObject("adodb.connection")
	conn.Open Application("Ba_jxqy_connstr")
	set rst=server.CreateObject("adodb.recordset")
	rst.Open "select * from ��Ʒ where ������='"&username&"' and ����='ҩƷ' and ����>0 and ����='"&cur&"'",conn
	if rst.EOF or rst.BOF then
		msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>����ҩ��</FONT>���޷�ʹ��ʹ����"&cur&"<br>');"
	else
		hpadd=rst("����")
		mpadd=rst("����")
		conn.Execute "update �û� set ����=����+"&hpadd&",����=����+"&mpadd&" where ����='"&username&"'"
		conn.Execute "update ��Ʒ set ����=����-1 where ������='"&username&"' and ����='"&cur&"'"
		msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>����ҩ��</FONT>��ʹ����"&cur&",��������"&hpadd&",��������"&mpadd&"<br>');parent.confrm.document.form1.hp.value='������"&hp+hpadd&"';parent.confrm.document.form1.mp.value='������"&mp+mpadd&"';parent.optfrm.location.replace('curative.asp');"
	end if
	rst.Close
	set rst=nothing
	conn.Close
	set conn=nothing	
end if	
Response.Write "<script language=javascript>"&msg&"</script>"
%>	
