<!--#include file=conn.asp-->
<!--#include file="inc/const.asp" -->
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	dim title,currentPage,idlist
 	const MaxPerPage=25
 	title=request("txtitle")
 	if not isempty(request("page")) then
		currentPage=cint(request("page"))
 	else
		currentPage=1
 	end if
	if not isempty(request("delAnnounce")) then
		idlist=request("delAnnounce")
		if instr(idlist,",")>0 then
			dim idarr
			idArr=split(idlist)
			dim id
			for i = 0 to ubound(idarr)
				id=clng(idarr(i))
				if request("agree")="批准" then
					call agreeannounce(clng(id))
					response.write "批 准 成 功！<br>"
					response.end
				end if
				if request("noagree")="不批准" then
					call noagreeannounce(clng(id))
					response.write " 拒 绝 成 功！<br>"
					response.end
				end if
			next
		else
			if request("agree")="批准" then
				call agreeannounce(clng(idlist))
				response.write "批 准 成 功！<br>"
				response.end
			end if
			if request("noagree")="不批准" then
				call noagreeannounce(clng(idlist))
				response.write " 拒 绝 成 功！<br>"
				response.end
			end if
		end if
	end if%>
	<table border="0" width="95%" cellspacing="0" cellpadding="0" align="center">
		<tr>
			<td width="100%" align="center"><form method=Post action="?"><%
				Set rs= Server.CreateObject("ADODB.Recordset")
				sql="select * from [user] where vip=4 order by userid desc"
				rs.open sql,conn,1,1
				dim totalPut,pagecount,curcount
			 	if rs.eof and rs.bof then
			 		response.write "还 没 有 任 何 申 请"
			 	else
			 		totalPut=rs.recordcount
					if totalPut mod MaxPerPage=0 then
						pagecount=totalPut \ MaxPerPage
					else
						pagecount=totalPut \ MaxPerPage+1
					end if
					if currentpage > pagecount then currentpage = pagecount
					if currentpage < 1 then currentpage=1
					rs.PageSize = MaxPerPage
					rs.AbsolutePage=currentpage
					curcount=0
					%><table class="tableBorder" cellspacing="1" width="95%" cellpadding="0" align="center">
 						<tr>
 							<th width="30%" align="center" height="20" >用户名</td>
							<th width="30%" align="center">注册时间</td>
							<th width="40%" align="center"><input type='submit' name="agree" class=buttonface value='批准'>&nbsp; <input type='submit' name="noagree" class=buttonface value='不批准'>全选：<input type=checkbox value="on" name="chkall" onclick="CheckAll(this.form)"> </th>
						</tr>
						<%do while not rs.eof
							curcount=curcount+1
							%><tr>
								<td height="23" width="30%" class=forumrow align="center"><a href=DISPUSER.ASP?name=<%=rs("username")%> target=_blank><%=rs("username")%></a></td>
								<td width="30%" class=forumrow align="center"><%=rs("addDate")%></td>
								<td width="40%" class=forumrow align="center"><input type='checkbox' name='delAnnounce' value='<%=cstr(rs("UserID"))%>'></td>
							</tr>
							<%rs.movenext
						loop%>
						<tr>
							<td height="23" colspan=3 align="right" class=forumrow>分页：<%call DispPageNum(currentpage,pagecount,"?page=","")%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
						</tr>
					</table>
				<%end if
			%></td>
		</tr>
	</form></table>
<%end if%>
</body>

<script language="javascript">
<!--
function CheckAll(form) {
 for (var i=0;i<form.elements.length;i++) {
 var e = form.elements[i];
 if (e.name != 'chkall') e.checked = form.chkall.checked; 
 }
 }
//-->
</script>

<%sub noagreeannounce(id)
	dim rss,rs2,sql2,rs,sql
	set rss=conn.execute("select * from [user] where userid="&cstr(id)&"")
	set rs2=server.createobject("adodb.recordset") 
	sql2="select * from [message]" 
	rs2.open sql2,conn,1,3 
	rs2.addnew 
	rs2("sender")=forum_info(0) 
	rs2("incept")=rss("username") 
	rs2("title")="您的VIP会员申请没有得到批准！" 
	rs2("content")="您的VIP会员申请没有得到批准！请与管理员联系！" 
	rs2("flag")=0 
	rs2("issend")=1 
	rs2.update 
	rs2.close 
	rss.close
	conn.execute("update [user] set vip=0 where userid="&cstr(id)&"")
End sub
 
sub agreeannounce(id)
	dim rs2,sql2,user
	user=conn.execute("select username from [user] where userid="&cstr(id)&"")(0)
	conn.execute("update [user] set vip=1,vipdate=now(),userwealth2=userwealth,userep2=userep,usercp2=usercp,userpower2=userpower,article2=article where userid="&cstr(id))
	set rs2=server.createobject("adodb.recordset") 
	sql2="select * from [message]" 
	rs2.open sql2,conn,1,3 
	rs2.addnew 
	rs2("sender")=forum_info(0) 
	rs2("incept")=user 
	rs2("title")="您的VIP会员申请已经得到批准！" 
	rs2("content")="您的VIP会员申请已经得到批准！您现在已经是正式的VIP会员了！" 
	rs2("flag")=0 
	rs2("issend")=1 
	rs2.update 
	rs2.close
End sub%>
