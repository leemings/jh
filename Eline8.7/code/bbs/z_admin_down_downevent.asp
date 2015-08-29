<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<!--#include file="z_down_function.asp" -->
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
<%dim MaxPerPage,title,currentpage,idlist
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	MaxPerPage=eventnum
	title=request("txtitle")
	if not isempty(request("page")) then
		currentPage=cint(request("page"))
	else
		currentPage=1
	end if
	if not isempty(request("selAnnounce")) then
		idlist=request("selAnnounce")
		sql="delete from downevent where id in ("&idlist&")"
		conndown.execute sql
		response.write "操作成功！"
		response.end
	end if%>
	<table border="0" width="95%" cellspacing="0" cellpadding="0" align=center>
		<form method=Post action="?">
		<tr>
			<td width="100%"><a href="z_admin_down_adminevent.asp">管理员管理日志</a>　　<font color=<%=forum_body(8)%>>付费用户下载日志</font>　　<a href="z_admin_down_userevent.asp">付费用户管理日志</a></td>
		</tr>
		<tr>
			<td width="100%" align=center><font color="#FF0000">付费用户下载日志</font></td>
		<tr>
		<tr>
			<td width="100%" align=center>
				<table border="0" cellspacing="1" cellpadding="3" width="100%" class=tableborder>
					<tr>
					  <th width="15%" height="25">ID</th>
					  <th width="55%">事件</td>
					  <th width="15%">时间</td>
					  <th width="15%">操作</td>
					</tr>
					<%dim n,totalput
					sql="select * from downevent order by id desc"
					Set rs= Server.CreateObject("ADODB.Recordset")
					rs.open sql,conndown,1,1
					if rs.eof and rs.bof then
						response.write "<tr><td height=25 class=forumrow align=center colspan=4>还 没 有 任 何 事 件</td></tr>"
					else
						totalPut=rs.recordcount
						if totalPut mod MaxPerPage=0 then
							n= totalPut \ MaxPerPage
						else
							n= totalPut \ MaxPerPage+1
						end if
						if currentpage > n then currentpage = n
						if currentpage < 1 then currentpage = 1
						rs.PageSize = MaxPerPage
						rs.AbsolutePage=currentpage
						i=0
						do while not rs.eof and i<MaxPerPage
							i=i+1%>
							<tr>
								<td height="25" width="15%" class=forumrow align=center><%=rs("id")%></td>
								<td width="55%" class=forumrow>&nbsp;<%=rs("username")&"下载《"&rs("buydown")&"》花费（"&rs("buypoint")&"元现金）"%></td>
								<td width="15%" class=forumrow align=center><%=rs("addtime")%></td>
								<td width="15%" class=forumrow align=center><input type='checkbox' name='selAnnounce' value='<%=cstr(rs("ID"))%>'></td>
							</tr>
							<%rs.movenext
						loop%>
						<tr>
							<td height="25" class=forumrow align=right colspan=4>页次：<b><%=currentPage%></b> / <b><%=n%></b> 页 每页 <b><%=maxperpage%></b> 事件数 <b><%=totalPut%></b>&nbsp;&nbsp;&nbsp;&nbsp;分页：<%call disppagenum(currentpage,n,"?page=","")%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
						</tr>
						<tr>
							<td height="25" class=forumrow align=center colspan=4><input type='submit' class=buttonface value='删 除'>&nbsp;&nbsp;全选：<input type=checkbox value="on" name="chkall" onclick="CheckAll(this.form)"></td>
						</tr>
					<%end if
					rs.close
					set rs=nothing%>
				</table>
			</td>
		</tr>
		</form>
	</table>
<%end if%>
</body>
<script language="javascript">
function CheckAll(form)  {
	for (var i=0;i<form.elements.length;i++) {
		var e = form.elements[i];
		if (e.name == 'selAnnounce') e.checked = !e.checked; 
	}
}
</script>
