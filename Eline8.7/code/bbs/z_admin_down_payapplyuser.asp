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
<%dim MaxPerPage,title,idlist,currentpage
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
	if not isempty(request("delAnnounce")) then
		idlist=request("delAnnounce")
		if instr(idlist,",")>0 then
			dim idarr
			idArr=split(idlist,",")
			dim id
			for i = 0 to ubound(idarr)
				id=clng(idarr(i))
				if request("agree")="批准" then
					call agreeannounce(id)
					response.write "批 准 成 功！<br>"
				end if
				if request("noagree")="不批准" then
					call noagreeannounce(id)
					response.write " 拒 绝 成 功！<br>"
				end if
			next
			response.end
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
				dim pagecount,totalput,curcount
				sql="select * from [user] where apply=2 order by id desc"
				Set rs= Server.CreateObject("ADODB.Recordset")
				rs.open sql,conndown,1,1
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
					curcount=0%>
					<table class="tableBorder" cellspacing="1" width="95%" cellpadding="0" align="center">
 						<tr>
 							<th width="30%" align="center" height="20" >用户名</td>
							<th width="30%" align="center">申请时间</td>
							<th width="40%" align="center"><input type='submit' name="agree" class=buttonface value='批准'>&nbsp; <input type='submit' name="noagree" class=buttonface value='不批准'>全选：<input type=checkbox value="on" name="chkall" onclick="CheckAll(this.form)"> </th>
						</tr>
						<%do while not rs.eof and curcount<rs.pagesize
							curcount=curcount+1
							%><tr>
								<td height="23" width="30%" class=forumrow align="center"><a href=DISPUSER.ASP?name=<%=rs("username")%> target=_blank><%=rs("username")%></a></td>
								<td width="30%" class=forumrow align="center"><%=rs("regtime")%></td>
								<td width="40%" class=forumrow align="center"><input type='checkbox' name='delAnnounce' value='<%=cstr(rs("ID"))%>'></td>
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
function CheckAll(form) {
 for (var i=0;i<form.elements.length;i++) {
 var e = form.elements[i];
 if (e.name == 'delAnnounce') e.checked = !e.checked; 
 }
 }
</script>

<%sub noagreeannounce(id)
	dim rss
	set rss=conndown.execute("select * from [user] where id="&cstr(id)&"")
	set rs=server.createobject("adodb.recordset")      
	sql="select * from [message]"      
	rs.open sql,conn,1,3      
	rs.addnew      
	rs("sender")=forum_info(0)&"下载中心"      
	rs("incept")=rss("username")     
	rs("title")="您的付费下载用户申请没有得到批准！"      
	rs("content")="您的付费下载用户申请没有得到批准！请与管理员联系！"      
	rs("flag")=0      
	rs("issend")=1      
	rs.update      
	rs.close   
	set rs=nothing
	rss.close
	set rss=nothing
	sql="delete from [user] where id="&cstr(id)
	conndown.execute sql
End sub
  
sub agreeannounce(id)
	dim rss,mess
	set rss=conndown.execute("select * from [user] where id="&cstr(id)&"")
	sql="update [user] set apply=1,regtime=date() where id="&cstr(id)
	conndown.execute sql
	set rs=server.createobject("adodb.recordset")
	sql="select * from events"
	rs.open sql,conndown,1,3
	mess=""&membername&"批准添加了付费用户“"&rss("username")&"”"	
	rs.addnew
	rs("event")=mess
	rs("addtime")=date()
	rs.update
	rs.close
	sql="select * from [message]"      
	rs.open sql,conn,1,3      
	rs.addnew      
	rs("sender")=forum_info(0)&"下载中心"      
	rs("incept")=rss("username")
	rs("title")="您的付费下载用户申请已经得到批准！"      
	rs("content")="您的付费下载用户申请已经得到批准！您现在已经是正式的付费用户了！"      
	rs("flag")=0      
	rs("issend")=1      
	rs.update      
	rs.close
	sql="select * from [userevent]"
	rs.open sql,conndown,1,3
	rs.addnew
	rs("username")=rss("username")
	rs("userevent")="您的付费下载用户申请已经得到批准！"
	rs("addtime")=date()
	rs("conuser")=membername
	rs.update
	rs.close  	
	set rs=nothing
	rss.close
	set rss=nothing
End sub%>
