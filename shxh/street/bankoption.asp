<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
nowdate=cstr(date())
bkmn=Request.Form("money")
bkop=Request.Form("op")
if isnumeric(bkmn) then
	if bkmn>1 then
		if bkop="ȡ��" then	sqlstr="update �û� set ����=����+"&bkmn&",���=���-"&bkmn&" where ����='"&username&"' and ���>="&bkmn
		if bkop="���" then sqlstr="update �û� set ����=����-"&bkmn&",���=���+"&bkmn&" where ����='"&username&"' and ����>="&bkmn
		if sqlstr="" then
			msg="���룿�������룿���Բ��𣿣�"
		else	
			set conn=server.CreateObject("adodb.connection")
			conn.Mode=16
			conn.IsolationLevel=256
			conn.Open Application("Ba_jxqy_connstr")
			conn.Execute(sqlstr)
			conn.Close
			set conn=nothing
			msg="������ɣ������գ�"
		end if	
	else
		msg="����Ц�İɣ��㵽������滹����ȡ��"
	end if	
else	
	msg="����Ǯѽ���Ҳ�̫����"
end if	
%>
<head><link rel="stylesheet" href="../style.css"></head>
<title><%=Application("Ba_jxqy_systemname")%></title>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color=0000FF size=5><%=Application("Ba_jxqy_systemname")%></font><p>
<%=msg%>
<p align=center><a href="bank.asp" onmouseover="window.status='����Ǯׯ';return true;" onmouseout="window.status='';return true;">����</a></p>
</body>