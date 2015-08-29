<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
</head>
<body leftmargin="0" topmargin="0">
<%dim menu
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	menu=request("menu")
	if menu=1 then
		call autoadmin()
	else
		call main()
	end if
end if
conndown.Close
set conndown=Nothing%>
</body>

<%sub main
	Set rs=Server.CreateObject("Adodb.RecordSet")   
	sql="select * from admin where flag>0 order by flag,id"   
	rs.Open sql,conndown,1,1%>
	<table border="0" width="95%" cellspacing="1" align=center>
		<tr>
			<td><font color="#FF0000">超级管理员</font>可以进行&nbsp;◇添加上传软件&nbsp;&nbsp;&nbsp;&nbsp;◇添加连接软件&nbsp;&nbsp;&nbsp;◇修改删除<br><font color="#FF0000">特殊管理员</font>可以进行&nbsp;◇添加上传软件&nbsp;&nbsp;&nbsp;&nbsp;◇添加连接软件&nbsp;&nbsp;&nbsp;&nbsp;<br><font color="#FF0000">普通管理员</font>可以进行&nbsp;◇添加连接软件<br><font color="#0000FF">自动添加管理员：</font>将您论坛的所有管理员加为主管理员，所有的超级版主加为超级管理员，版主和贵宾加为特殊管理员<br><font color="#0000FF">手工增加管理员：</font>自行增加各级别管理员</td>
		</tr>
		<tr>
			<td><br><a href=?menu=1><font color="#0000FF">自动添加管理员</font></a> | <a href=z_admin_down_adduser.asp><font color="#0000FF">手工添加管理员</font></a></td>
		</tr>
	</table>
	<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
		<tr height="27">
			<th width="5%">序号</th>
			<th width="15%">用户名</th>
			<th width="15%">加入时间</th>
			<th width="50%">权限</th>
			<th width="15%">操作</th>
		</tr>
		<%do while not rs.EOF %>
			<form method="post" action="z_admin_down_saveuser.asp">
			<tr height="30">
				<td width="5%" align=center class=forumrow><%=rs("id")%></td>
				<td width="15%" align=center class=forumrow><%=rs("username")%></td>
				<td width="15%" align=center class=forumrow><%=rs("addtime")%></td>
				<td width="50%" align=center class=forumrow><%
					if rs("flag")>1 then
						%><input type=radio name="flag" value="2" <%if rs("flag")=2 then%>checked<%end if%>>超级管理员&nbsp;<input type=radio name="flag" value="3" <%if rs("flag")=3 then%>checked<%end if%>>特殊管理员&nbsp;<input type=radio name="flag" value="4" <%if rs("flag")=4 then%>checked<%end if%>>普通管理员<%
					else
						response.write "下载中心主管理员"
					end if
				%></td>
				<td width="15%" align=center class=forumrow><%
					if rs("flag")<=1 and lcase(trim(rs("username")))<>lcase(trim(membername)) then
						%><input type="submit" name="Submit" value="删除"><input type="hidden" name="user" value="<%=rs("username")%>"><%
					elseif rs("flag")>1 then
						%><input type="submit" name="Submit" value="修改"> <input type="submit" name="Submit" value="删除"><input type="hidden" name="user" value="<%=rs("username")%>"><%
					end if
				%></td>
			</tr>
			</form>
		<%rs.MoveNext
	loop
	rs.Close
	set rs=Nothing%>
	</table>
	<br>
<%end sub

sub autoadmin()
	dim usergroupid1,usergroupid2,usergroupid3,usergroupid4,rss
	dim rs1,sql1
	usergroupid1=1
	usergroupid2=2
	usergroupid3=3
	usergroupid4=8
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	sql1="select * from [admin]"
	rs1.open sql1,conndown,1,3
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from [user] where usergroupid="&usergroupid1
	rs.open sql,conn,1,1
	if not(rs.eof and rs.bof) then
		do while not rs.eof
			set rss=conndown.execute("select * from [admin] where username='"&rs("username")&"'")
			if rss.eof and rss.bof then
				rs1.addnew
				rs1("username")=rs("username")
				rs1("flag")=1
				rs1("addtime")=date()
				rs1.update
			end if
			rss.close
			rs.movenext
		loop
	end if
	rs.close
	sql="select * from [user] where usergroupid="&usergroupid2
	rs.open sql,conn,1,1
	if not(rs.eof and rs.bof) then
		do while not rs.eof
			set rss=conndown.execute("select * from [admin] where username='"&rs("username")&"'")
			if rss.eof and rss.bof then
				rs1.addnew
				rs1("username")=rs("username")
				rs1("flag")=2
				rs1("addtime")=date()
				rs1.update
			end if
			rss.close
			rs.movenext
		loop
	end if
	rs.close
	sql="select * from [user] where usergroupid="&usergroupid3&" or usergroupid="&usergroupid4
	rs.open sql,conn,1,1
	if not(rs.eof and rs.bof) then
		do while not rs.eof
			set rss=conndown.execute("select * from [admin] where username='"&rs("username")&"'")
			if rss.eof and rss.bof then
				rs1.addnew
				rs1("username")=rs("username")
				rs1("flag")=3
				rs1("addtime")=date()
				rs1.update
			end if
			rss.close
			rs.movenext
		loop
	end if
	rs.close
	set rs=nothing
	rs1.close
	set rs1=nothing
	response.write "修改成功！"
end sub
%>
