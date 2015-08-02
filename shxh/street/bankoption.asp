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
		if bkop="取款" then	sqlstr="update 用户 set 银两=银两+"&bkmn&",存款=存款-"&bkmn&" where 姓名='"&username&"' and 存款>="&bkmn
		if bkop="存款" then sqlstr="update 用户 set 银两=银两-"&bkmn&",存款=存款+"&bkmn&" where 姓名='"&username&"' and 银两>="&bkmn
		if sqlstr="" then
			msg="你想？？，我想？？对不起？？"
		else	
			set conn=server.CreateObject("adodb.connection")
			conn.Mode=16
			conn.IsolationLevel=256
			conn.Open Application("Ba_jxqy_connstr")
			conn.Execute(sqlstr)
			conn.Close
			set conn=nothing
			msg="操作完成，请验收！"
		end if	
	else
		msg="开玩笑的吧，你到底是想存还是想取！"
	end if	
else	
	msg="多少钱呀，我不太明白"
end if	
%>
<head><link rel="stylesheet" href="../style.css"></head>
<title><%=Application("Ba_jxqy_systemname")%></title>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color=0000FF size=5><%=Application("Ba_jxqy_systemname")%></font><p>
<%=msg%>
<p align=center><a href="bank.asp" onmouseover="window.status='返回钱庄';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>