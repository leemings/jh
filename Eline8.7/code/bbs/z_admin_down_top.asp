<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<base target="top">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else%>
	<br>
	&nbsp;◇<a href="z_admin_down_softadd.asp" target="footer" title="请准备好上传文件"><b>添加上传软件</b></a>&nbsp;
	&nbsp;◇<a href="z_admin_down_freeadd.asp" target="footer" title="直接添加下载地址等信息"><b>添加连接软件</b></a>&nbsp; 
	&nbsp;◇<a href="z_admin_down_adminedit.asp" target="footer"><b>修改删除</b></a>&nbsp;  
	&nbsp;◇<a href="z_admin_down_class.asp" target="footer"><b>栏目管理</b></a>&nbsp; 
	&nbsp;◇<a href="z_admin_down_unite.asp" target="footer"><b>辅助功能</b></a>&nbsp; 
	&nbsp;◇<a href="z_admin_down_adminuser.asp" target="footer"><b>管理员</b></a>&nbsp;
	&nbsp;◇<a href="z_admin_down_adminevent.asp" target="footer"><b>查看事件</b></a>&nbsp;
	&nbsp;◇<a href="z_admin_down_paymanage.asp" target="footer"><b>付费管理</b></a>
<%end if%>
</body>
</html>