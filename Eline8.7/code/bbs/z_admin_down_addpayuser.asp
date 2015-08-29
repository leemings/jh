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
<%dim menu
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	menu=request("menu")
	if menu=1 or menu="" then
		call main()
	elseif menu=2 then
		call addpayuser()
	end if
end if

sub main()%>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form method="post" action="?menu=2" name="form1"  onsubmit="javascript:return check();">
		<tr>
			<th colspan="2">增加新的付费用户</th>
		</tr>
		<tr>
			<td class=forumrow colspan="2">注意事项： 请确认此用户是本论坛的用户，如果是多个用户请用逗号分隔开</td>
		</tr>
		<tr>
			<td class=forumrow width="30%">用户名：</td>
			<td class=forumrow width="70%"><input maxLength=20 name=username size=20></td>
		</tr>
		<tr>
			<td class=forumrow colspan=2 align=center><input type="submit" name="Submit" value="确定"></td>
		</tr>
		</form>	
	</table>
<%end sub

sub addpayuser()
	dim user,rs_f,mess
	if request("username")="" then	
		response.write "您忘记填写付费用户了吧"
		exit sub
	else
		user=CheckStr(request("username"))
		user=split(user,",")
	end if
	for i=0 to ubound(user)
		set rs_f=conndown.execute("select * from [user] where username='"&user(i)&"'")
		if not(rs_f.eof and rs_f.bof) then
			Response.Write "用户"&user(i)&"已经是付费用户了<br>"
		else
			rs_f.close
			set rs_f=conn.execute("select * from [user] where username='"&user(i)&"'")
			if rs_f.eof and rs_f.bof then
				Response.Write "用户"&user(i)&"不是论坛用户<br>"
			else
				sql="select * from [user]"
				Set rs= Server.CreateObject("ADODB.Recordset")
				rs.open sql,conndown,1,3
				rs.addnew
				rs("username")=user(i)
				rs("allpoint")=0
				rs("regtime")=date()
				rs("lockuser")=1
				rs("apply")=1
				rs.update
				rs.close
				sql="select * from events"
				rs.open sql,conndown,1,3
				mess=""&membername&"添加了付费用户“"&user(i)&"”"	
				rs.addnew
				rs("event")=mess
				rs("addtime")=date()
				rs.update
				rs.close
				sql="select * from [message]"      
				rs.open sql,conn,1,3      
				rs.addnew      
				rs("sender")=forum_info(0)&"下载中心"      
				rs("incept")=user(i)
				rs("title")="您的付费下载用户账号已经开通！"      
				rs("content")="您的付费下载用户账号已经开通！请妥善保管您的密码！"      
				rs("flag")=0      
				rs("issend")=1      
				rs.update      
				rs.close
				response.write "新的付费用户"&user(i)&"添加成功！<br>"
			end if
		end if
	next
end sub%>