<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else%>
	<frameset rows="40,*" frameborder="yes" border="1" framespacing="2">
		<noframes>
			<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
			<p>此网页使用了框架，但您的浏览器不支持框架。</p>
  		</body>
		</noframes>
		<frame name="top" target="footer" scrolling="no" marginwidth="0" marginheight="0" src="z_admin_down_top.asp" noresize>
		<frame name="footer" scrolling="auto" marginwidth="0" marginheight="0" src="z_admin_down_main.asp">
	</frameset>
<%end if%>