<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="z_plus_check.asp"-->
<%dim menu
menu=request.querystring("menu")
stats="VIP会员申请"
call nav()
call head_var(2,0,"","")
if membername="" or not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有VIP会员申请的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
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
call activeonline()
call footer() 

sub main()
	dim sql1,rs1,vipmod
	sql="select * from [vip]"
	set rs=conn.execute(sql)
	if not rs.eof then
		vipmod=rs("vipmod")
	end if
	sql1="select * from [user] where username='"&membername&"'"
	set rs1=conn.execute(sql1)
	if rs1.eof and rs1.bof then
		Errmsg=Errmsg+"<br>"+"<li>您不是论坛的合法用户！"
		exit sub
	end if%>
	<table cellspacing=1 align=center class=tableborder1 width="97%" height="228">
	<tr width="100%"><th valign=middle align=center height=21 width="100%" colspan="5">申请说明</td></tr>
	<tr width="100%">
	<td valign=middle align=left height=21 width="100%" class=tablebody1 colspan="5"><blockquote><%
		if vipmod=0 then
			%><br><font color=red>本会员系统当前采用的是自动添加VIP会员模式</font><br>如果您当前的条件满足VIP会员申请条件，您就可以直接申请成为<%=Forum_info(0)%>的VIP会员<%
		elseif vipmod=1 then
			%><br><font color=red>本会员系统当前采用的是人工添加VIP会员模式</font><br>如果您当前的条件满足VIP会员申请条件，您就可以申请成为<%=Forum_info(0)%>的VIP会员，<br>您的申请会被提交给管理员审批，一但通过审批即可成为VIP会员<%
		elseif vipmod=2 then
			%><br><font color=red>本会员系统当前采用的是半自动添加VIP会员模式</font><br>如果您当前的条件满足VIP会员首次申请条件，您就可以申请首次成为<%=Forum_info(0)%>的VIP会员，<br>您的申请会被提交给管理员审批，一但通过审批即可成为VIP会员；<br>如果您当前的条件满足再次申请条件，您就可以直接申请成为<%=Forum_info(0)%>的再次VIP会员<%
		end if
	%><br>1、VIP会员的有效期为 <b><%=rs("vipdate")%></b> 天<br>2、当您的有效期快到期时，我们会提前用短信形式通知您<br>3、谢谢您对<%=Forum_info(0)%>的支持<br></blockquote></td>
	</tr>
	<tr width="100%">
		<td valign=middle align=center height=21 width="100%" class=tablebody1 colspan="5">欢迎<B><%=membername%></B>申请VIP会员<br></td>
	</tr>
	<tr width="100%">
		<th valign=middle align=center height=24 width="20%"><b>首次所需要条件</b></th>
		<th valign=middle align=center height=24 width="20%"><b>首次VIP会员条件</b></th>
		<th valign=middle align=center height=24 width="20%"><b>再次VIP会员条件</b></th>
		<th valign=middle align=center height=24 width="20%"><b>您当前的条件</b></th>
		<th valign=middle align=center height=24 width="20%"><b>是否满足</b></th>
	</tr>
	<tr width="100%">
		<td valign=middle align=center height=21 width="20%" class=tablebody1>金钱</td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("wealth")<=0 then
				response.write "-"
			else
				response.write rs("wealth")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("wealth2")<=0 then
				response.write "-"
			else
				response.write rs("wealth2")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("wealth2")<=0 then
					response.write "-"
				else
					response.write rs1("userwealth")-rs1("userwealth2")
				end if
			else
				if rs("wealth")<=0 then
					response.write "-"
				else
					response.write rs1("userwealth")
				end if
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("wealth2")<=0 then
					response.write "-"
				else
					response.write iiif(rs("wealth2"),rs1("userwealth")-rs1("userwealth2"),"满足","<font color=red>不满足</font>")
				end if
			else	
				if rs("wealth")<=0 then
					response.write "-"
				else
					response.write iiif(rs("wealth"),rs1("userwealth"),"满足","<font color=red>不满足</font>")
				end if
			end if
		%></td>
	</tr>
	<tr width="100%">
		<td valign=middle align=center height=21 width="20%" class=tablebody1>帖子数</td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("article")<=0 then
				response.write "-"
			else
				response.write rs("article")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("article2")<=0 then
				response.write "-"
			else
				response.write rs("article2")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("article2")<=0 then
					response.write "-"
				else
					response.write rs1("article")-rs1("article2")
				end if
			else
				if rs("article")<=0 then
					response.write "-"
				else
					response.write rs1("article")
				end if
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("article2")<=0 then
					response.write "-"
				else
					response.write iiif(rs("article2"),rs1("article")-rs1("article2"),"满足","<font color=red>不满足</font>")
				end if
			else	
				if rs("article")<=0 then
					response.write "-"
				else
					response.write iiif(rs("article"),rs1("article"),"满足","<font color=red>不满足</font>")
				end if
			end if
		%></td>
	</tr>
	<tr width="100%">
		<td valign=middle align=center height=21 width="20%" class=tablebody1>经验值</td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("userep")<=0 then
				response.write "-"
			else
				response.write rs("userep")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("userep2")<=0 then
				response.write "-"
			else
				response.write rs("userep2")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("userep2")<=0 then
					response.write "-"
				else
					response.write rs1("userep")-rs1("userep2")
				end if
			else
				if rs("userep")<=0 then
					response.write "-"
				else
					response.write rs1("userep")
				end if
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("userep2")<=0 then
					response.write "-"
				else
					response.write iiif(rs("userep2"),rs1("userep")-rs1("userep2"),"满足","<font color=red>不满足</font>")
				end if
			else	
				if rs("userep")<=0 then
					response.write "-"
				else
					response.write iiif(rs("userep"),rs1("userep"),"满足","<font color=red>不满足</font>")
				end if
			end if
		%></td>
	</tr>
	<tr width="100%">
		<td valign=middle align=center height=21 width="20%" class=tablebody1>魅力值</td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("usercp")<=0 then
				response.write "-"
			else
				response.write rs("usercp")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("usercp2")<=0 then
				response.write "-"
			else
				response.write rs("usercp2")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("usercp2")<=0 then
					response.write "-"
				else
					response.write rs1("usercp")-rs1("usercp2")
				end if
			else
				if rs("usercp")<=0 then
					response.write "-"
				else
					response.write rs1("usercp")
				end if
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("usercp2")<=0 then
					response.write "-"
				else
					response.write iiif(rs("usercp2"),rs1("usercp")-rs1("usercp2"),"满足","<font color=red>不满足</font>")
				end if
			else	
				if rs("usercp")<=0 then
					response.write "-"
				else
					response.write iiif(rs("usercp"),rs1("usercp"),"满足","<font color=red>不满足</font>")
				end if
			end if
		%></td>
	</tr>
	<tr width="100%">
		<td valign=middle align=center height=21 width="20%" class=tablebody1>威望值</td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("power")<=0 then
				response.write "-"
			else
				response.write rs("power")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("power2")<=0 then
				response.write "-"
			else
				response.write rs("power2")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("power2")<=0 then
					response.write "-"
				else
					response.write rs1("userpower")-rs1("userpower2")
				end if
			else
				if rs("power")<=0 then
					response.write "-"
				else
					response.write rs1("userpower")
				end if
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("power2")<=0 then
					response.write "-"
				else
					response.write iiif(rs("power2"),rs1("userpower")-rs1("userpower2"),"满足","<font color=red>不满足</font>")
				end if
			else	
				if rs("power")<=0 then
					response.write "-"
				else
					response.write iiif(rs("power"),rs1("userpower"),"满足","<font color=red>不满足</font>")
				end if
			end if
		%></td>
	</tr>
	<tr width="100%">
	<td valign=middle align=center width="25%" class=tablebody1><a href="toplist.asp?orders=9"><font color="#FF0000">查询VIP会员</font></a></td>
	<td valign=middle align=center width="50%" colspan="3" class=tablebody1><%
		if rs1("vip")=2 then
			response.write "<br><font color=red>您被限制，不能成为"&Forum_info(0)&"VIP会员</font><br><br>"
		elseif rs1("vip")=3 then
			response.write "<br><font color=red>您已经是VIP会员，但您的VIP权限已经被锁定，请不要重复申请！</font><br><br>"
		elseif rs1("vip")=4 then
			response.write "<br><font color=red>您已经递交过申请，尚未得到批准，请不要重复申请！</font><br><br>"
		elseif rs1("vip")=1 then
			response.write "<br><font color=red>您已经是"&Forum_info(0)&"VIP会员</font><br>"
			response.write "<br><font color=red>注销您的会员资格？</font><br>"
			%><form name="form1" method="post" action="z_vip.asp?menu=2"><INPUT name=username size="20" value="<%=membername%>" type=hidden><INPUT type=submit value=注销 name=submit></form><%
		else
			if not isnull(rs1("vipdate")) then
				if (rs("wealth2")>rs1("userwealth")-rs1("userwealth2") and rs("wealth2")>0) or (rs("article2")>rs1("article")-rs1("article2") and rs("article2")>0) or (rs("userep2")>rs1("userep")-rs1("userep2") and rs("userep2")>0) or (rs("usercp2")>rs1("usercp")-rs1("usercp2") and rs("usercp2")>0) or (rs("power2")>rs1("userpower")-rs1("userpower2") and rs("power2")>0) then
					response.write "<br><br><font color=red>您至少有一项条件不满足，您暂时还无法再次申请成为"&Forum_info(0)&"VIP会员</font><br><br>"
				else
					response.write "<br><font color=red>您可以申请VIP会员</font><br>"
					%><form name="form2" method="post" action="z_vip.asp?menu=1"><INPUT name=username size="20" value="<%=membername%>" type=hidden><INPUT type=submit value=申请 name=submit></form><%
				end if
			else
				if (rs("wealth")>rs1("userwealth") and rs("wealth")>0) or (rs("article")>rs1("article") and rs("article")>0) or (rs("userep")>rs1("userep") and rs("userep")>0) or (rs("usercp")>rs1("usercp") and rs("usercp")>0) or (rs("power")>rs1("userpower") and rs("power")>0) then
					response.write "<br><br><font color=red>您至少有一项条件不满足，您暂时还无法首次申请成为"&Forum_info(0)&"VIP会员</font><br><br>"
				else
					response.write "<br><font color=red>您可以申请VIP会员</font><br>"
					%><form name="form2" method="post" action="z_vip.asp?menu=1"><INPUT name=username size="20" value="<%=membername%>" type=hidden><INPUT type=submit value=申请 name=submit></form><%
				end if
			end if
		end if
	%></td>
	<td valign=middle align=center width="25%" class=tablebody1><a href="toplist.asp?orders=10"><font color="#FF0000">等待VIP审批会员</font></a><br><br><%
		if rs1("vip")=1 then
			response.write "您上次成为VIP会员的时间为：<br>"&formatdatetime(rs1("vipdate"),2)&"<br>"
			response.write "离您的有效期还有：<b>"&cint(rs("vipdate")-datediff("d",rs1("vipdate"),now()))&"</b>天"
		end if
	%></td>
	</tr>
	</table>
	<%rs.close
	Set rs = Nothing
	rs1.close
	Set rs1 = Nothing
end sub

sub apply()
	dim boarduser
	dim boarduser_1
	dim userlen
	dim updateinfo
	dim rs1,sql1,rs,sql,err
	sql="select * from [vip]"
	set rs=conn.execute(sql)
	sql1="select * from [user] where username='"&membername&"'"
	set rs1=conn.execute(sql1)
	if rs1("vip")=1 then
		Errmsg=Errmsg+"<br>"+"<li>您已经是"&Forum_info(0)&"的VIP会员，请不要重复申请！"
		call dvbbs_error()
	elseif rs1("vip")=2 then
		Errmsg=Errmsg+"<br>"+"<li>您被限制，不能成为"&Forum_info(0)&"VIP会员！"
		call dvbbs_error()
	elseif rs1("vip")=3 then
		Errmsg=Errmsg+"<br>"+"<li>您已经是VIP会员，但您的VIP权限已经被锁定，请不要重复申请！"
		call dvbbs_error()
	elseif rs1("vip")=4 then
		Errmsg=Errmsg+"<br>"+"<li>您已经递交过申请，尚未得到批准，请不要重复申请！"
		call dvbbs_error()
	else
		dim sucword
		sucword=""
		if not isnull(rs1("vipdate")) then
			if (rs("wealth2")>rs1("userwealth")-rs1("userwealth2") and rs("wealth2")>0) or (rs("article2")>rs1("article")-rs1("article2") and rs("article2")>0) or (rs("userep2")>rs1("userep")-rs1("userep2") and rs("userep2")>0) or (rs("usercp2")>rs1("usercp")-rs1("usercp2") and rs("usercp2")>0) or (rs("power2")>rs1("userpower")-rs1("userpower2") and rs("power2")>0) then
				Errmsg=Errmsg+"<br>"+"<li>您至少有一项条件不满足，您暂时还无法再次申请成为"&Forum_info(0)&"的VIP会员"
				call dvbbs_error()
			else
				if rs("vipmod")=0 or rs("vipmod")=2 then
					conn.execute("update [user] set vip=1,vipdate=now(),userwealth2=userwealth,userep2=userep,usercp2=usercp,userpower2=userpower,article2=article where username='"&membername&"'")
					sucword="您申请成功，恭喜您再次成为"&Forum_info(0)&"的VIP会员"
				elseif rs("vipmod")=1 then
					conn.execute("update [user] set vip=4 where username='"&membername&"'")
					sucword="您的申请已经提交，请等待管理员的审批！"
				end if
			end if
		else
			if (rs("wealth")>rs1("userwealth") and rs("wealth")>0) or (rs("article")>rs1("article") and rs("article")>0) or (rs("userep")>rs1("userep") and rs("userep")>0) or (rs("usercp")>rs1("usercp") and rs("usercp")>0) or (rs("power")>rs1("userpower") and rs("power")>0) then
				Errmsg=Errmsg+"<br>"+"<li>您至少有一项条件不满足，您暂时还无法首次申请成为"&Forum_info(0)&"的VIP会员"
				call dvbbs_error()
			else
				if rs("vipmod")=0 then
					conn.execute("update [user] set vip=1,vipdate=now(),userwealth2=userwealth,userep2=userep,usercp2=usercp,userpower2=userpower,article2=article where username='"&membername&"'")
					sucword="您申请成功，恭喜您首次成为"&Forum_info(0)&"的VIP会员"
				elseif rs("vipmod")=1 or rs("vipmod")=2 then
					conn.execute("update [user] set vip=4 where username='"&membername&"'")
					sucword="您的申请已经提交，请等待管理员的审批！"
				end if
			end if
		end if
		if sucword<>"" then
			%><table cellspacing=1 align=center class=tableborder1 width="97%" >
				<tr width="100%"><th valign=middle align=center height=21 width="100%">首次申请VIP会员</tr>
				<tr width="100%">
					<td valign=middle align=center height=21 width="100%" class=tablebody1><br><font color=red><%=sucword%></font><br><br><a href=index.asp>返回首页</a><br></td>
				</tr>
			</table><%
		end if
	end if
	rs1.close
	set rs1=nothing
	rs.close
	set rs=nothing
end sub

sub cancel()
	dim boarduser
	dim boarduser_1
	dim userlen
	dim updateinfo
	dim rs1,sql1,rs,sql
	sql="select * from [vip]"
	set rs=conn.execute(sql)
	conn.execute("update [user] set vip=0 where username='"&membername&"'")
	%><table cellspacing=1 align=center class=tableborder1 width="97%" >
		<tr width="100%"><th valign=middle align=center height=21 width="100%">注销VIP会员</tr>
		<tr width="100%">
			<td valign=middle align=center height=21 width="100%" class=tablebody1><br><font color=red>您注销成功，欢迎您成为<%=Forum_info(0)%>的普通会员</font><br><br><a href=index.asp>返回首页</a><br></td>
		</tr>
		</table><%
	rs.close
	set rs=nothing
end sub%>



