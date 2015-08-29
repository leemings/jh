<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else%>
<table border="0" cellspacing="1" width="80%" align=center>
	<tr>
		<td width="100%"><br><p align=center><b>下载中心管理页面</b></p><b>管理员可进行操作说明</b>：<br>
			<br>1、通过Web上传供下载软件，请选择添加上传软件以及准备好上传文件。<br><br>
			<br>2、已经上传文件至服务器，请选择添加连接软件进行添加。<br>
			<br>3、对已经添加软件修改或删除，请点左边相关连接进行操作。<br>
			<br>4、对软件栏目进行添加修改删除，请点左边相关连接进行操作。<br>
			<br>5、下载属性分为：<font color="#800080">开放下载</font>（所有访客及会员都可以下载），<font color="#800080">会员下载</font>（仅限论坛会员下载），<font color="#800080">VIP下载</font>（仅限论坛VIP会员下载），<font color="#800080">特约下载</font>（仅限版主和贵宾，超级版主和管理员下载）<br>
			<br>6、下载系统管理员说明：<font color="#800080">主论坛管理员</font>（可以管理所有项目），<font color="#800080">超级管理员</font>（可以进行上传，增加软件，软件的删除和修改），<font color="#800080">特殊管理员</font>（可以进行上传，增加软件），<font color="#800080">普通用户</font>（增加软件）<br>
		</td>
	</tr>
</table>
<%end if%>