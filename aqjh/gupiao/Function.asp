<%
function iif(expression,returntrue,returnfalse)
	if expression=0 then
		iif=returnfalse
	else
		iif=returntrue
	end if
end function

'取两个数中最大那个
function Max(expression1,expression2)
	if expression1+0>=expression2+0 then
		Max=expression1
	else
		Max=expression2
	end if
end function

Rem 过滤HTML代码
function HTMLEncode(fString)
	if not isnull(fString) then
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
	
		fString = Replace(fString, CHR(32), "&nbsp;")
		fString = Replace(fString, CHR(9), "&nbsp;")
		fString = Replace(fString, CHR(34), "&quot;")
		fString = Replace(fString, CHR(39), "&#39;")
		fString = Replace(fString, CHR(13), "")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10), "<BR> ")
	
		HTMLEncode = fString
	end if
end function
'-------------------------------信息提示-------------------------------
'flag=0 显示关闭页面	flag=1 显示返回股票交易大厅		flag=2	显示返回上一页
sub endinfo(flag) 
%>
<br>
<table  width="75%" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR><TD BACKGROUND="Images/topbg.gif" height=9></TD></TR> 
	<tr><td valign=center align=middle height=23 background="Images/Header.gif"><b>股票操作信息提示</b></td>
	<tr><td height=1 bgcolor="#E4E8EF">
<%
	if errmess<>"" then
		response.Write "<b>产生错误的可能原因：</b><br><br><li>您是否仔细阅读了股市指南，可能您还没有登陆或者不具有使用当前功能的权限"	
		response.Write "<font color=red>"&errmess&"</font>"
	else
		response.Write "<b>操作成功：</b><br><br><li>欢迎光临快乐江湖股票交易中心，请返回进行其它操作"
		response.Write "<font color=navy>"&sucmess&"</font>"
	end if		
%>	
	<br></td></tr> 
	<TR><TD BACKGROUND="Images/title.gif" height=21 align="center">
<%
	if flag=0 then
		response.Write "<a href=""javascript:window.close()"">[离开股票交易中心]</a><a href=""#"" onclick=""window.close()"">[关闭本页]</a>"
	else
		if rUrl="" or isnull(rurl) then
			rUrl=Request.ServerVariables("HTTP_REFERER")
		end if
		response.Write "<a href=""stock.asp"" class=cblue>[返回股市大厅]</a>&nbsp;&nbsp;"	
		response.Write "<a href="""&rUrl&""" class=cblue>[返回上一页]</a>"
	end if
%></TD></TR> 
</table>
<%
end sub

sub AdminHead()
%>
<br>
<table cellspacing=1 cellpadding=3 align=center width="97%" bgcolor="#0066CC">
	<tr>
		<td align=left valign=middle background="Images/title.gif" height="21"> &nbsp;<a href="stock.asp">股票交易大厅</a> ｜ <a href="Admin_Setting.asp">基本设置</a> ｜ <a href="Admin_Gupiao.asp">股票管理</a> ｜ <a href="Admin_User.asp">客户管理</a> ｜ <a href="Announcements.asp">公告管理</a> ｜ <a href="Admin_Data.asp">数据库管理</a> ｜ <a href=javascript:history.go(-1)>返回上一页</a></td>
	</tr> 
</table>
<br>
<%
end sub
'----------------------------------------
Rem 在线股民
sub GPOnline()
	dim statuserid
	statuserid=replace(Request.ServerVariables("REMOTE_HOST"),".","")
	if membername="" or MyUserID=0 then
		Dim TempUsername
		TempUsername=iif(membername="","客人",membername)
		session("GP_UserID")=statuserid
		sql="select id from online where id="&cstr(session("GP_UserID"))
		set rs=conn.execute(sql)
		if rs.eof and rs.bof then
			sql="insert into online(id,username,Uid,ip,startime,lastimebk,stats,userhidden) values ("&statuserid&",'"&TempUsername&"',0,'"&Request.ServerVariables("REMOTE_HOST")&"',Now(),Now(),'"&replace(stats,"'","")&"',2)"
		else
			sql="update online set lastimebk=Now(),stats='"&replace(stats,"'","")&"' where id="&cstr(session("GP_UserID"))
		end if
		conn.execute(sql)
	else
		sql="select id from online where Uid="&MyUserID
		set rs=conn.execute(sql)
		if rs.eof and rs.bof then
			sql="insert into online(id,username,Uid,ip,startime,lastimebk,stats,userhidden) values ("&statuserid&",'"&membername&"',"&MyUserID&",'"&Request.ServerVariables("REMOTE_HOST")&"',Now(),Now(),'"&replace(stats,"'","")&"',"&userhidden&")"		
		else
			sql="update online set lastimebk=Now(),stats='"&replace(stats,"'","")&"' where Uid="&MyUserID
		end if
		conn.execute(sql)
		conn.execute("update [客户] set 最后日期=now() where id="&MyUserID)
		rs.close
		if session("GP_UserID")<>"" then
			Conn.Execute("delete from online where id="&session("GP_UserID"))
			session("GP_UserID")=""
		end if
	end if
	set rs=nothing
	Rem 删除超时用户
	sql="Delete FROM online WHERE DATEDIFF('s', lastimebk, now()) > "&Gupiao_Setting(9)&"*60"
	Conn.Execute sql
end sub
'在线人数
function online(boardid)
	online=conn.execute("Select count(*) from online")(0)
	if isnull(online) then online=0
end function 
'显示在线股民列表
sub onlineuser(online_u,online_g)
response.write "<tr><td colspan=2><table cellpadding=0 cellspacing=0 border=0 width=""100%"" style=""word-break:break-all;""><tr height=20 valign=middle>"
dim userip,NowStats
if cint(online_u)=1 then		'在线显示会员信息
	i=0
	'用户信息
	sql="select username,ip,stats,userhidden,Uid,startime,lastimebk from online where uid>=0"
	set rs=conn.execute(sql)
	do while not rs.eof
	NowStats="目前位置：" & server.HTMLEncode(rs(2)) &"<br>"
	if master then
		userip="真实ＩＰ：" & rs(1)
	else
		userip="真实ＩＰ：已设置保密"
	end if
	
	if membername=rs(0) then		'自己
		response.write "<td background=""images/title.gif"" width=""14%"">&nbsp;<a href=dispu.asp?uid="&rs(4)&" target=_blank class=cblue title="""& NowStats & UserIP &""">"&server.htmlencode(rs(0))&"</font></a></td>"
	else
		if rs(3)=1 then			'如果是隐身股民
			if master then
				response.write "<td background=""images/title.gif"" width=""14%"">&nbsp;<a href=dispu.asp?uid="&rs(4)&" target=_blank title="""& NowStats & UserIP & "<br>当前状态：隐身"">"&server.htmlencode(rs(0))&"</a></td>"
			else
				response.write "<td background=""images/title.gif"" width=""14%"">&nbsp;<a href=# class=cgray target=_blank title="""& NowStats & UserIP &""">隐身股民</a></td>"
			end if
		else
			response.write "<td background=""images/title.gif"" width=""14%"">&nbsp;<a href=dispu.asp?uid="&rs(4)&" target=_blank title="""& NowStats & UserIP &""">"&server.htmlencode(rs(0))&"</a></td>"
		end if
	end if
	if i=6 then response.write "</tr><tr height=20 valign=middle>"
	if i>6 then 
		i=1
	else
		i=i+1
	end if
	rs.movenext
	loop
end if
response.write "</tr></TABLE>"
set rs=nothing
end sub
'-------------------------------股票收市了-------------------------------
sub CloseGuPiao(expression)
	if expression=1 then
%>
	<title><%=Gupiao_Setting(5)%>-股市已经收市了</title>
	<style type=text/css><!--
		td {  font-family: 宋体; font-size: 9pt}-->
	</style>
	<br>
	<table cellspacing=1 cellpadding=3 align=center width="75%" bgcolor="#0066CC" border=0>
	<TR><TD BACKGROUND="Images/topbg.gif" height=9 colspan=3></TD></TR> 
	<tr><td height=23 background="Images/header.gif" align="center">股票收市通知</td></tr>
	<tr><td width="100%" bgcolor="#FFFFFF" height="50" valign="middle">
						亲爱的股民 <font color=blue><%=membername%></font>，您好！<br>
						&nbsp;&nbsp;&nbsp;&nbsp;现在股票市场已经收市了，请于每天的 <%=split(Gupiao_Setting(4),"||")(0)%>:00～<%=split(Gupiao_Setting(4),"||")(1)%>:00 再来。谢谢合作！ 
	</td></tr>
	<tr><td align=center height=26 bgcolor="#E4E8EF"><a href="javascript:window.close()">[离开股票交易中心]</a><a href="#" onClick="window.close()">[关闭本页]</a></td></tr> 
	</table>
	<%elseif expression=0 then%>
	<title><%=Gupiao_Setting(5)%>-股市暂时关闭</title>
	<style type=text/css><!--
		td {  font-family: 宋体; font-size: 9pt}--> 
	</style>
	<br><br>
	<table cellspacing=1 cellpadding=3 align=center width="75%" bgcolor="#0066CC" border=0>
		<TR><TD BACKGROUND="Images/topbg.gif" height=9 colspan=3></TD></TR> 
		<tr><td height=23 background="Images/header.gif" align="center">欢迎光临 <%=Gupiao_Setting(5)%></td></tr>
		<tr>
			<td bgcolor="#FFFFFF"><br><font color="#FF3366"> □ 现在股市正处于关闭状态中，原因如下：</font><br><br>&nbsp;&nbsp;&nbsp;<%=StopReadme%><br><br></td>
		</tr>
		<tr><td align=center height=26 bgcolor="#E4E8EF"><a href="javascript:window.close()">[离开股市交易中心]</a><a href="#" onClick="window.close()">[关闭本页]</a></td></tr> 
	</table>
	<%else%>
	<title><%=Gupiao_Setting(5)%>-防刷新功能开启</title>
	<style type=text/css><!--
		td {  font-family: 宋体; font-size: 9pt}--> 
	</style> 
	<META http-equiv=Content-Type content=text/html; charset=gb2312> 
	<br><br>
	<table cellspacing=1 cellpadding=3 align=center width="75%" bgcolor="#0066CC" border=0>
		<TR><TD BACKGROUND="Images/topbg.gif" height=9 colspan=3></TD></TR> 
		<tr><td height=23 background="Images/header.gif" align="center">防刷新机制</td></tr>
		<tr>
			<td bgcolor="#FFFFFF">
				<br>&nbsp;&nbsp;<font color="#FF3366">本页面起用了防刷新机制，请不要在 <font color=blue><%=Gupiao_Setting(7)%></font> 秒内连续刷新本页面</font><br><br> 
				&nbsp;&nbsp;<font color=blue><%=Gupiao_Setting(7)%></font> 秒之后将会自动打开页面，请稍后……<br><br>
			</td>  
		</tr> 
		<tr><td align=center height=26 bgcolor="#E4E8EF" id="TdReFlash"></td></tr> 
	</table>
	<meta HTTP-EQUIV=REFRESH CONTENT="<%=Gupiao_Setting(7)%>">
	<script language="VBScript"> 
	<!--
		TimeLimit=<%=Gupiao_Setting(7)%>
		call GetSec()
		function GetSec()
			TimeLimit=TimeLimit-1
			TdReFlash.innerHTML = "<font color=blue>"&TimeLimit&"</font>"
			if TimeLimit<= 0 then
				TdReFlash.innerHTML="<font color=red>正在打开页面……</font>"
			else
				setTimeout "GetSec()",1000 
			end if
		end function
	//-->	
	</script>		
<%
	end if
end sub 

sub History()
	dim rst,TongJi
	set rst=conn.execute("select TongJi,当前价格,sid,剩余股份,IniTradeNum from [股票] order by sid")
		
	do while not rst.eof
		if rst(3)>rst(4) then
			sql="交易量=IniTradeNum "
		elseif rst(3)>0 then
			sql="交易量=剩余股份 "
		else
			sql="交易量=0,状态='封' "
		end if	
		TongJi=rst(0)
		if TongJi="" or isnull(TongJi) then TongJi="0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
		TongJi=mid(TongJi,instr(TongJi,"|")+1)&"|"&formatnumber(rst(1),2,-1)
		conn.execute("update [股票] set TongJi='"&TongJi&"',开盘价格=当前价格,成交量=0,买入笔数=0,卖出笔数=0,TodayWave=0,"&sql&" where sid="&rst(2))
		rst.movenext
	loop
end sub


%>