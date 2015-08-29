<!--#include file=conn.asp-->
<!--#include file="inc/const.asp" -->
<%dim menu,body
menu=request.querystring("menu")%>
<html>
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
	if menu="" then
 		call main()
	elseif menu=1 then
 	call manage()
	elseif menu=2 then
 	call addman()
 elseif menu=3 then
 	call savemod()
 elseif menu=4 then
 	call cleanvip()
 end if 	
end if%>

<%sub main() 
dim sql1,rs1,sql2,rs2 
sql1="select * from [vip]" 
set rs1=conn.execute(sql1) 
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
<form action=?menu=3 method=post>
<tr><th valign=middle height=23 align=left width="100%">VIP会员申请系统模式设定</th></tr>
<tr><td valign=middle class=forumrow align=left width="75%">
说明：VIP会员申请系统有三种模式<br>
<font color="#FF0000">自动添加系统</font>：用户在前台可以看到申请条件，如果满足条件即可申请，申请后直接加为VIP会员；VIP到期后如果满足再次申请条件直接加为VIP会员。<br>
<font color="#FF0000">人工添加系统</font>：用户在前台可以看到申请条件，如果满足条件可以递交申请，申请会提交给管理员审批，如果批准则加为VIP会员；VIP到期后如果满足条件可以递交再次申请，申请会提交给管理员审批，如果批准则再次加为VIP会员。<br> 
<font color="#FF0000">半自动添加系统</font>：用户在前台可以看到申请条件，如果满足条件即可申请，申请后直接加为VIP会员；VIP到期后如果满足再次申请条件直接加为VIP会员。</td>
</tr><tr>
<td valign=middle class=forumrow align=left width="75%">
<p align="center">自动添加VIP会员<input type=radio name="vipmod" value="0" <%if rs1("vipmod")=0 then%>checked<%end if%>>　　　　人工添加VIP会员<input type=radio name="vipmod" value="1" <%if rs1("vipmod")=1 then%>checked<%end if%>>　　　　半自动添加VIP会员<input type=radio name="vipmod" value="2" <%if rs1("vipmod")=2 then%>checked<%end if%>>　　　<input type=submit name="submit" value="提 交"></td></tr>
</form>
</table>
<br>

<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
<tr><th align=center colspan=4 width="95%" height=1><p align="left"><b>论坛VIP申请条件管理</b></p></th></tr> 
<form name="form" method="post" action="?menu=1"> 
<tr> 
	<td align=right width="170" height="25" class=forumrow>首次申请VIP所需金额：</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userwealth size="20" value="<%=rs1("wealth")%>"></td> 
	<td align=right width="170" height="25" class=forumrow>再次申请VIP所需金额：</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userwealth2 size="20" value="<%=rs1("wealth2")%>"></td> 
</tr> 
<tr> 
	<td align=right width="170" height="25" class=forumrow>首次申请VIP所需经验：</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userep size="20" value="<%=rs1("userep")%>"></td> 
	<td align=right width="170" height="25" class=forumrow>再次申请VIP所需经验：</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userep2 size="20" value="<%=rs1("userep2")%>"></td> 
</tr> 
<tr> 
	<td align=right width="170" height="25" class=forumrow>首次申请VIP所需魅力：</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=usercp size="20" value="<%=rs1("usercp")%>"></td> 
	<td align=right width="170" height="25" class=forumrow>再次申请VIP所需魅力：</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=usercp2 size="20" value="<%=rs1("usercp2")%>"></td> 
</tr> 
<tr> 
	<td align=right width="170" height="25" class=forumrow>首次申请VIP所需威望：</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userpower size="20" value="<%=rs1("power")%>"></td> 
	<td align=right width="170" height="25" class=forumrow>再次申请VIP所需威望：</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userpower2 size="20" value="<%=rs1("power2")%>"></td> 
</tr> 
<tr> 
	<td align=right width="170" height="25" class=forumrow>首次申请VIP所需文章数：</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userart size="20" value="<%=rs1("article")%>"></td> 
	<td align=right width="170" height="25" class=forumrow>再次申请VIP所需文章数：</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userart2 size="20" value="<%=rs1("article2")%>"></td> 
</tr> 
<tr> 
	<td align=right width="170" height="25" class=forumrow>VIP会员有效期：(天)</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=vipdate size="20" value="<%=rs1("vipdate")%>"></td> 
	<td align=right width="170" height="25" class=forumrow>到期提前通知的天数：(天)</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=notifydays size="20" value="<%=rs1("notifydays")%>"></td> 
</tr> 
<tr> 
	<td align=center colspan=4 width="95%" height="27" class=forumrow><INPUT type=submit value=修改 name=submit></td> 
</tr> 
</form> 
</table> 
<br> 
<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder"> 
<tr><th align=center width="95%" height=1 colspan="2"><p align="left">论坛VIP会员管理</p></th></tr>
<tr width="95%" > 
<td width="35%" class=forumrow><b>增加VIP会员</b>（请输入您想确认为VIP的用户名，请确认此用户是本论坛的用户）</td>
<form name="form1" method="post" action="?menu=2"> 
<td width="65%" class=forumrow height="45">如果是多个用户请用逗号分隔开<br><INPUT name=vipname size="20">&nbsp; <INPUT type=submit value=确认 name=submit></td> 
</form> 
</tr> 
<tr width="95%" > 
<td width="35%" class=forumrow><b>VIP会员管理</b></td>
<form name="form1" method="get" action="z_admin_vipuser.asp">
<td width="65%" class=forumrow height="40"><INPUT type=submit value='进 行 管 理' name=submit></td> 
</form>
</tr> 
<tr width="95%" > 
<td width="35%" class=forumrow><b>审批VIP会员申请</b></td> 
<form name="form1" method="get" action="z_admin_vipapplyuser.asp">
<td width="65%" class=forumrow height="40"><INPUT type=submit value='进 行 审 批' name=submit></td> 
</form> 
</tr> 
<tr width="95%" > 
<td width="35%" class=forumrow><b>清理已过期的VIP会员</b></td> 
<form name="form1" method="get" action="?menu=4">
<td width="65%" class=forumrow height="40"><INPUT type=submit value='进 行 清 理' name=submit></td> 
</form> 
</tr> 
</table> 
<br> 
<%end sub

sub savemod()
	dim rs,sql
	if request("vipmod")="" then
		response.write "选择项不能为空！"
		response.end
	else
		set rs= server.createobject ("adodb.recordset") 
		sql = "select * from vip" 
		rs.Open sql,conn,1,3 
		if not(rs.eof and rs.bof) then 
			rs("vipmod")=request("vipmod") 
			rs.Update 
		end if
		rs.close
		set rs=nothing 
		response.write"修改成功！" 
	end if
end sub

sub manage() 
	dim money,usercp,userep,userpower,userart,rs,sql,vipdate,money2,usercp2,userep2,userpower2,userart2,notifydays
	if not isInteger(request.form("userwealth")) or not isInteger(request.form("userart")) or not isInteger(request.form("userep")) or not isInteger(request.form("usercp")) or not isInteger(request.form("userpower")) or not isInteger(request.form("vipdate")) or not isInteger(request.form("userwealth2")) or not isInteger(request.form("userart2")) or not isInteger(request.form("userep2")) or not isInteger(request.form("usercp2")) or not isInteger(request.form("userpower2")) or not isInteger(request.form("notifydays")) then
		response.write "输入的数据必须为整数！"
		response.end 
	else 
		if request.form("userwealth")="" then 
			money=20
		else 
			money=int(request.form("userwealth")) 
		end if 
		if request.form("userep")="" then 
			userep=20 
		else 
			userep=int(request.form("userep")) 
		end if 
		if request.form("usercp")="" then 
			usercp=20 
		else 
			usercp=int(request.form("usercp")) 
		end if 
		if request.form("userpower")="" then 
			userpower=20 
		else 
			userpower=int(request.form("userpower")) 
		end if 
		if request.form("userart")="" then 
			userart=20 
		else 
			userart=int(request.form("userart")) 
		end if 
		if request.form("vipdate")="" then 
			vipdate=90 
		else 
			vipdate=int(request.form("vipdate")) 
		end if 
		if request.form("userwealth2")="" then 
			money2=20
		else 
			money2=int(request.form("userwealth2")) 
		end if 
		if request.form("userep2")="" then 
			userep2=20 
		else 
			userep2=int(request.form("userep2")) 
		end if 
		if request.form("usercp2")="" then 
			usercp2=20 
		else 
			usercp2=int(request.form("usercp2")) 
		end if 
		if request.form("userpower2")="" then 
			userpower2=20 
		else 
			userpower2=int(request.form("userpower2")) 
		end if 
		if request.form("userart2")="" then 
			userart2=20 
		else 
			userart2=int(request.form("userart2")) 
		end if 
		if request.form("notifydays")="" then 
			notidydays=5 
		else 
			notifydays=int(request.form("notifydays")) 
		end if 
	end if 
	set rs= server.createobject ("adodb.recordset") 
	sql = "select * from vip" 
	rs.Open sql,conn,1,3 
	if not(rs.eof and rs.bof) then 
		rs("wealth") = money 
		rs("userep") = userep 
		rs("usercp") = usercp 
		rs("power") = userpower 
		rs("article") = userart 
		rs("vipdate")=vipdate           
		rs("wealth2") = money2 
		rs("userep2") = userep2 
		rs("usercp2") = usercp2 
		rs("power2") = userpower2 
		rs("article2") = userart2 
		rs("notifydays") = notifydays 
		rs.Update 
	end if 
	rs.close 
	set rs=nothing 
	response.write "修改成功!" 
end sub 

sub addman() 
	dim rs1,sql1,rs,sql,incept,vipdate
	dim msgcontent
	if request("vipname")="" then	
		response.write "您忘记填写VIP对象了吧"
		exit sub
	else
		incept=CheckStr(request("vipname"))
		incept=split(incept,",")
	end if
	sql="select * from [vip]"     
	set rs1=conn.execute(sql)
	vipdate=rs1("vipdate")
	rs1.close
	for i=0 to ubound(incept)
		sql="select vip from [user] where username='"&replace(incept(i),"'","")&"'"
		set rs=conn.execute(sql)
		if rs.eof and rs.bof then
			response.write "论坛没有"&replace(incept(i),"'","")&"这个用户，看看你的VIP对象写对了嘛？<BR>"
		else
			if isnull(rs(0)) or rs(0)=4 or rs(0)=0 then
				conn.execute("update [user] set vip=1,vipdate=now(),userwealth2=userwealth,userep2=userep,usercp2=usercp,userpower2=userpower,article2=article where username='"&replace(incept(i),"'","")&"'")
				msgcontent="您的VIP资格已经生效，谢谢您对"& forum_info(0) &"的支持！"&CHR(10)&"您的会员有效期为"&vipdate&"天,请您在有效期内好好使用VIP会员的权力！"
				conn.execute("insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&forum_info(0)&"','您已正式成为VIP会员！','"&msgcontent&"',Now(),0,1)")
				response.write "添加成功！<BR>"
			end if
		end if
		rs.close
	next
end sub

sub cleanvip()
	dim vipdate,msgcontent,vipmod,money2,userep2,usercp2,power2,article2
	set rs=conn.execute("select vipdate,vipmod,wealth2,userep2,usercp2,power2,article2 from [vip]")
	vipdate=rs(0)
	vipmod=rs(1)
	money2=rs(2)
	userep2=rs(3)
	usercp2=rs(4)
	power2=rs(5)
	article2=rs(6)
	rs.close
	set rs=conn.execute("select username,vipdate,userwealth,userep,usercp,userpower,article,userwealth2,userep2,usercp2,userpower2,article2 from [User] where vip=1 or vip=3")
	do while not rs.eof
		if datediff("d",rs(1),now)>vipdate then
			if vipmod<>1 then
				if (rs("userwealth")-rs("userwealth2")>=money2 or money2<=0) and (rs("userep")-rs("userep2")>=userep2 or userep2<=0) and (rs("usercp")-rs("usercp2")>=usercp2 or usercp2<=0) and (rs("userpower")-rs("userpower2")>=power2 or power2<=0) and (rs("article")-rs("article2")>=article2 or article2<=0) then
					conn.execute ("update [user] set vipdate=now(),userwealth2=userwealth,userep2=userep,usercp2=usercp,userpower2=userpower,article2=article where username='"&rs(0)&"'")
					msgcontent="您的VIP资格到期，系统已经将您自动加为本期VIP会员，谢谢您对"& forum_info(0) &"的支持！"
					conn.execute("insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&rs(0)&"','"&forum_info(0)&"','您的VIP会员资格续期成功！','"&msgcontent&"',Now(),0,1)")
				else
					conn.execute ("update [User] set vip=0 where username='"&rs(0)&"'")
					msgcontent="您的VIP资格已经过期，您现在尚未满足再次申请的条件，谢谢您对"& forum_info(0) &"的支持！"&CHR(10)&"如果您愿意，可以在满足条件后重新申请VIP会员！"
					conn.execute("insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&rs(0)&"','"&forum_info(0)&"','您的VIP会员资格续期失败！','"&msgcontent&"',Now(),0,1)")
				end if
			else
				conn.execute ("update [User] set vip=0,vipdate=now() where username='"&rs(0)&"'")
				msgcontent="您的VIP资格已经过期，谢谢您对"& forum_info(0) &"的支持！"&CHR(10)&"如果您愿意，您可以再次申请VIP会员！"
				conn.execute("insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&rs(0)&"','"&forum_info(0)&"','您的VIP会员资格已经过期！','"&msgcontent&"',Now(),0,1)")
			end if
		end if
		rs.movenext
	loop
	response.write "清期过期用户成功"
end sub%> 
</html>