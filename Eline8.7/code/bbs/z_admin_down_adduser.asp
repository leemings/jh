<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
<script language=javascript>
function check() {
	if(document.form1.username.value=="") {
		alert("用户名为空");
		return false;
	}
}
</script>
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
<%dim username,right_class
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	if request("submit")<>"" then
		username=Request("username")
		right_class=CInt(Request("right_class"))
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.Open "Select * from [user] where username='"&username&"'",conn,1,1
		if  rs.EOF and rs.bof then
			Response.Write "此用户不是论坛用户"
			rs.Close
			Set rs=Nothing
		else
			rs.close
			rs.Open "Select * from Admin where username='"&username&"'",conndown,1,1
			if rs.EOF and rs.bof then
				rs.Close
				set rs=Nothing
				sql="insert into Admin (username,Flag,addtime) values ('"&username&"',"&right_class&",date())"
				conndown.Execute sql
				conndown.Close
				set conndown=Nothing
				Response.Redirect "z_admin_down_adminuser.asp"
			else
				Response.Write "该用户名已经存在"
				rs.Close
				Set rs=Nothing
			end if
		end if
	else%>
		<table border="0" width="95%" cellspacing="1" align=center>
			<tr>
				<td><font color="#FF0000">超级管理员</font>可以进行&nbsp;◇添加上传软件&nbsp;&nbsp;&nbsp;&nbsp;◇添加连接软件&nbsp;&nbsp;&nbsp;◇修改删除&nbsp;&nbsp;<br><font color="#FF0000">特殊管理员</font>可以进行&nbsp;◇添加上传软件&nbsp;&nbsp;&nbsp;&nbsp;◇添加连接软件&nbsp;&nbsp;&nbsp;&nbsp;<br><font color="#FF0000">普通管理员</font>可以进行&nbsp;◇添加连接软件</td>
			</tr>
		</table>
		<table width="95%" border="0" cellspacing="1" cellpadding="3" class=tableborder align=center>
		  <form method="post" action="?" name="form1" onsubmit="javascript:return check();">
			<tr> 
				<th height="25" align="center">新增管理员</th>
			</tr>
			<tr> 
				<td height="30" align="center" class=forumrow>用 户 名<input type="text" name="username">&nbsp; <font color="#FF0000">(请确定此用户为论坛成员)</font></td>
			</tr>
			<tr> 
				<td height="30" align="center" class=forumrow>权限设置</td>
			</tr>
			<tr>
				<td height="30" align="center" class=forumrow><input type="radio" name="right_class" value="4" checked>普通管理员&nbsp;&nbsp;<input type="radio" name="right_class" value="3">特殊管理员&nbsp;&nbsp;<input type="radio" name="right_class" value="2">超级管理员</td>
			</tr>
			<tr>
				<td height="30" align="center" class=forumrow><input type="submit" name="Submit" value="确定"></td>
			</tr>
			</form>
		</table>
	<%end if%>
<%end if
conndown.Close
Set conndown=Nothing%>
</body>
