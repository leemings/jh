<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="z_down_function.asp" -->
<%dim menu
menu=request.querystring("menu")
stats="付费用户申请"
call nav()
call head_var(0,0,"下载中心","z_down_default.asp")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有下载中心付费用户申请的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	call dvbbs_error()
else
	if menu="" then
 		call main()
	elseif menu=1 then
  	call apply()
	elseif menu=2 then
  	call cancel()
	end if  	
end if
conndown.close
set conndown=nothing
call activeonline()
call footer()

sub main()
	sql="select * from [payconfig]"
	set rs=conndown.execute(sql)%>
	<table cellspacing=1 align=center class=tableborder1>
		<tr>
			<th valign=middle align=center height=25 width="100%" colspan="4">申请说明</th>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="100%" class=tablebody1 colspan="4"><%
				if paymod=1 then
					%><br><font color=red>本下载系统当前采用的是自动添加付费用户模式</font><br>如果您当前的条件满足付费用户申请条件，您就可以直接申请成为下载中心的付费用户<br><br><%
				else
					%><br><font color=red>本下载系统当前采用的是人工添加付费用户模式</font><br>如果您当前的条件满足付费用户申请条件，您就可以申请成为下载中心的付费用户，<br>您的申请会被提交给管理员审批，一旦通过审批即可成为付费用户<br><br><%
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="100%" class=tablebody1 colspan="4">欢迎<B><%=membername%></B>申请下载中心付费用户<br></td>
		</tr>
		<tr>
			<th valign=middle align=center height=25 width="25%"><b>所需要条件</b></th>
			<th valign=middle align=center width="25%">付费用户<b>条件</b></th>
			<th valign=middle align=center width="25%"><b>您当前的条件</b></th>
			<th valign=middle align=center width="25%"><b>是否满足</b></th>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="25%" class=tablebody1>金钱</td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("wealth")<=0 then
					response.write "-"
				else
					response.write rs("wealth")
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("wealth")<=0 then
					response.write "-"
				else
					response.write mymoney
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("wealth")<=0 then
					response.write "-"
				else
					response.write iiif(rs("wealth"),mymoney,"满足","<font color=red>不满足</font>")
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="25%" class=tablebody1>帖子数</td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("article")<=0 then
					response.write "-"
				else
					response.write rs("article")
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("article")<=0 then
					response.write "-"
				else
					response.write myarticle
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("article")<=0 then
					response.write "-"
				else
					response.write iiif(rs("article"),myarticle,"满足","<font color=red>不满足</font>")
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="25%" class=tablebody1>经验值</td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("userep")<=0 then
					response.write "-"
				else
					response.write rs("userep")
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("userep")<=0 then
					response.write "-"
				else
					response.write myuserep
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("userep")<=0 then
					response.write "-"
				else
					response.write iiif(rs("userep"),myuserep,"满足","<font color=red>不满足</font>")
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="25%" class=tablebody1>魅力值</td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("usercp")<=0 then
					response.write "-"
				else
					response.write rs("usercp")
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("usercp")<=0 then
					response.write "-"
				else
					response.write myusercp
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("usercp")<=0 then
					response.write "-"
				else
					response.write iiif(rs("usercp"),myusercp,"满足","<font color=red>不满足</font>")
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="25%" class=tablebody1>威望值</td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("power")<=0 then
					response.write "-"
				else
					response.write rs("power")
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("power")<=0 then
					response.write "-"
				else
					response.write mypower
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("power")<=0 then
					response.write "-"
				else
					response.write iiif(rs("power"),mypower,"满足","<font color=red>不满足</font>")
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="25%" class=tablebody1>积分值</td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("userscore")<=0 then
					response.write "-"
				else
					response.write rs("userscore")
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("userscore")<=0 then
					response.write "-"
				else
					response.write myuserscore
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("userscore")<=0 then
					response.write "-"
				else
					response.write iiif(rs("userscore"),myuserscore,"满足","<font color=red>不满足</font>")
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center width="50%" colspan="4" class=tablebody1><%
				if rs("wealth")>mymoney or rs("article")>myarticle or rs("userep")>myuserep or rs("usercp")>myusercp or rs("power")>mypower or rs("userscore")>myuserscore then
					response.write "<br><font color="&forum_body(8)&">您至少有一项条件不满足，您暂时还无法申请成为"&Forum_info(0)&"下载中心付费会员</font><br>"
				else
					set rs=conndown.execute("select * from [user] where username='"&membername&"'")
					if rs.eof and rs.bof then
						response.write "<br><font color="&forum_body(8)&">您可以申请成为下载中心付费会员</font><br>"
						%><form name="form2" method="post" action="?menu=1"><INPUT type=submit value=申请 name=submit></form><%
					else
						if rs("apply")=2 then
							response.write "<br><font color="&forum_body(8)&">您的申请还未审批</font><br>"
						elseif rs("apply")=1 then
							response.write "<br><font color="&forum_body(8)&">您已是下载中心的付费会员</font><br>"
						end if
					end if
				end if
			%></td>
		</tr>
	</table>
<%end sub

sub apply()
	dim rss,errword
	sql="select * from [payconfig]"
	set rs=conndown.execute(sql)
	set rss = server.CreateObject ("Adodb.recordset")
	sql="select * from [user] where username='"&membername&"'"
	rss.open sql,conndown,1,3
	if rs("wealth")>mymoney or rs("article")>myarticle or rs("userep")>myuserep or rs("usercp")>myusercp or rs("power")>mypower or rs("userscore")>myuserscore then
		errword="您至少有一项条件不满足，您暂时还无法申请成为"&Forum_info(0)&"下载中心的付费用户"%>
		<table cellspacing=1 align=center class=tableborder1>
			<tr>
				<th valign=middle align=center height=25 width="100%">申请付费用户</th>
			</tr>
			<tr>
				<td valign=middle align=center height=25 width="100%" class=tablebody1><font color=red><%=errword%></font></td>
			</tr>
			<tr>
				<td valign=middle align=center height=25 width="100%" class=tablebody2><a href=z_down_default.asp>返回下载中心首页</a></td>
			</tr>
		</table>
	<%elseif not(rss.eof and rss.bof) then
		errword="您已递交过申请或是您已经是付费会员了，请不要重复申请"%>
		<table cellspacing=1 align=center class=tableborder1>
			<tr>
				<th valign=middle align=center height=25 width="100%">申请付费用户</th>
			</tr>
			<tr>
				<td valign=middle align=center height=25 width="100%" class=tablebody1><font color=<%=forum_body(8)%>><%=errword%></font></td>
			</tr>
			<tr>
				<td valign=middle align=center height=25 width="100%" class=tablebody2><a href=z_down_default.asp>返回下载中心首页</a></td>
			</tr>
		</table>
	<%else
		rss.addnew
		rss("username")=membername
		rss("regtime")=date()
		rss("allpoint")=0
		rss("lockuser")=1
		if paymod=1 then
			rss("apply")=1 
		else
			rss("apply")=2 
		end if
		rss.update%>
		<table cellspacing=1 align=center class=tableborder1>
			<tr>
				<th valign=middle align=center height=25 width="100%">申请付费用户</th>
			</tr>
			<tr>
				<td valign=middle align=center height=25 width="100%" class=tablebody1><font color=<%=forum_body(8)%>><%
					if paymod=1 then
						response.write "您的申请已经成功，恭喜您成为"&Forum_info(0)&"下载中心的付费用户"
					else
						response.write "您的申请已经递交，请等待管理员的审批，感谢您申请成为"&Forum_info(0)&"下载中心的付费用户"
					end if
				%></font></td>
			</tr>
			<tr>
				<td valign=middle align=center height=25 width="100%" class=tablebody2><a href=z_down_default.asp>返回下载中心首页</a></td>
			</tr>
		</table>
	<%end if
end sub%>
