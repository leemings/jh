<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%dim manager,userflag,submit
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	manager=LCase(Request("user"))
	userflag=Request("flag")
	submit=Trim(Request("submit"))
	if submit="�޸�" then
		sql="update Admin set flag="&userflag&" where username='"&manager&"'"   
		conndown.Execute sql
	elseif submit="ɾ��" then
		sql="delete from Admin where username='"&manager&"'"
		conndown.Execute sql
	end if
	conndown.Close
	set conndown=Nothing      
	response.redirect Request.ServerVariables("HTTP_REFERER")
end if%>
