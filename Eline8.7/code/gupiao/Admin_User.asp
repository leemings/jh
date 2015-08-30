<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head><title><%=Gupiao_Setting(5)%>-股市控制面板♀wWw.happyjh.com♀</title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<table  width="98%" border=0 cellspacing=1 cellpadding=0 align=center bgcolor="#0066CC">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
	</TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif"><font size="2"><b>股市管理中心</b></font></td>
	</tr>
	<tr><td bgcolor="#E4E8EF"> 
<% 
call AdminHead()
if not master then
	errmess="<li>本页为管理员专用，您没有管理本页的权限！"
	call endinfo(1)
else 
	select case request("action")
		case "EditUser"
			call EditUser()
		case "SaveEdit"		
			call SaveEdit()
		case "del"
			call DelUser()	
		case else
			call main()
	end select
end if
%>
<br>
</td></tr>
<tr><td height=32 background="images/footer.gif" valign=middle></td></tr>
</table>
</body></html>
<%
CloseDatabase		'关闭数据库
'=====================================
sub main()
	if request("action")="search" then
		dim username
		username=trim(replace(request("username"),"'",""))
		if username="" then
			errmess="<li>请输入查找关键字"
			rurl="?"
			call endinfo(2)
			exit sub
		end if	
		if request("usernamechk")="yes" then
			sql= "select * from 客户 where 帐号='"&username&"'"
		else	
			sql= "select * from 客户 where 帐号 like '%"&username&"%'"         
		end if
	else 
		sql= "select * from 客户 order by 资金 desc"	
	end if					
%>
<table border="0" width="97%" bgcolor="#0066CC" cellspacing="1" cellpadding="3" align="center">
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif" colspan="11"><b>客户管理</b></td>
	</tr>
	<tr bgcolor="#ffffff"> 
	<%if request("action")="search" then%>
	<td colspan="11" height=18><font color=red>搜索结果如下：</font><a href="?">[返回客户列表]</a></td>
	<%else%>	
	<form name="form1" method="post" action="?action=search">
		<td colspan="11" height=18>快速查找：<input type="text"  name=username> <input type="checkbox"  name="usernamechk" value="yes" checked>用户名完全匹配<input type="submit" name="Submit" value="查询客户"></td>
	</form>
	<%end if%>
	</tr>
	<tr bgcolor="#666666" align="center"> 
		<td background="Images/lan.gif"><font color=#ffffff>帐户</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>资金</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>总资金</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>持股种类</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>今日买入</font></td> 
		<td background="Images/lan.gif"><font color=#ffffff>今日卖出</font></td> 
		<td background="Images/lan.gif"><font color=#ffffff>总买入</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>总卖出</font></td> 
		<td background="Images/lan.gif"><font color=#ffffff>最后日期</font></td> 		
		<td background="Images/lan.gif"><font color=#ffffff>状态</font></td> 
		<td background="Images/lan.gif"><font color=#ffffff>管理</font></td>
	</tr>
<%
		dim currentpage,page_count,Pcount
		dim totalrec,endpage
		currentPage=request("page")
		if currentpage="" or not isnumeric(currentpage) then
			currentpage=1
		else
			currentpage=clng(currentpage)
			if err then
				currentpage=1
				err.clear
			end if
		end if
		Set rs= Server.CreateObject("ADODB.Recordset")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "<tr><td colspan=11 bgcolor=#E4E8EF> 没有找到任何股民</td></tr>"
		else
			rs.PageSize = 20 
			rs.AbsolutePage=currentpage
			page_count=0
			totalrec=rs.recordcount
			do while (not rs.eof) and (not page_count = 20)
%>
			<tr align=center> 
				<td bgcolor="#FFFFFF"><a href="?action=EditUser&username=<%=rs("帐号")%>"><%=rs("帐号")%></a></td>
				<td bgcolor="#FFFFFF" align="right"><%=formatnumber(rs("资金"),2,true)%></td>
				<td bgcolor="#FFFFFF" align="right"><%=formatnumber(rs("总资金"),2,true)%></td>
				<td bgcolor="#FFFFFF"><%=rs("持股种类")%></td>
				<td bgcolor="#FFFFFF"><%=rs("今日买入")%></td>
				<td bgcolor="#FFFFFF"><%=rs("今日卖出")%></td>
				<td bgcolor="#FFFFFF"><%=rs("总买入")%></td>
				<td bgcolor="#FFFFFF"><%=rs("总卖出")%></td>
				<td bgcolor="#FFFFFF"><%=formatdatetime(rs("最后日期"),2)%></td>				
				<td bgcolor="#FFFFFF"><%if rs("锁定")=1 then%><font color=red>锁定<%else%>正常<%end if%></td>
				<td bgcolor="#FFFFFF"><a href="?action=EditUser&username=<%=rs("帐号")%>">编辑</a> | <a href="?action=del&username=<%=rs("帐号")%>" onclick="javascript:{if(confirm('您确定删除 <%=rs("帐号")%> 这个客户吗?')){return true;}return false;}">删除</a></td>
			</tr>
<%			
				page_count = page_count + 1		
				rs.movenext
			loop
			Pcount=rs.PageCount
%>
			<tr bgcolor="#666666" align="center"> 
				<td background="Images/lan.gif"><font color=#ffffff>单位</font></td>
				<td background="Images/lan.gif" align="right"><font color=#ffffff>元</font></td>
				<td background="Images/lan.gif" align="right"><font color=#ffffff>元</font></td>
				<td background="Images/lan.gif"><font color=#ffffff>种</font></td>
				<td background="Images/lan.gif"><font color=#ffffff>笔</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>笔</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>笔</font></td>
				<td background="Images/lan.gif"><font color=#ffffff>笔</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>年-月-日</font></td> 		
				<td background="Images/lan.gif"><font color=#ffffff>-</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>-</font></td>
			</tr>
			<tr><td colspan=11 background="images/title.gif" align=right>
			<table border="0" cellpadding="0" cellspacing="0" width="100%"><tr>
			<td align=left>共有<font color=blue><%=totalrec%></font>个股民，每页<font color=blue><%=rs.PageSize%></font>个，第<font color=red><%=currentpage%></font>页/共<font color=blue><%=Pcount%></font>页</td>
			<td align=right>分页：
<%
			if currentpage > 4 then
				response.write "<a href=""?page=1"">[1]</a> ..."
			end if
			if Pcount>currentpage+3 then
			endpage=currentpage+3
			else
			endpage=Pcount
			end if
			for i=currentpage-3 to endpage
			if not i<1 then
				if i = clng(currentpage) then
				response.write " <font color=red>["&i&"]</font>"
				else
				response.write " <a href=""?page="&i&""">["&i&"]</a>"
				end if
			end if
			next
			if currentpage+3 < Pcount then 
				response.write "... <a href=""?page="&Pcount&""">["&Pcount&"]</a>"
			end if
			response.Write "</td></tr></table></td></tr>"		
		end if
		rs.close
		set rs=nothing	
%>
</table>
<% 
end sub
'---------------编辑用户---------
sub EditUser()
	dim username
	username=trim(replace(request("username"),"'",""))
	if username="" then
		errmess="<li>错误的参数，请指定用户名"
		call endinfo(2)
		exit sub
	end if	
	sql= "select * from 客户 where 帐号='"&username&"'"
	set rs=conn.execute(sql)         
	if rs.eof then
		errmess="<li>该用户的不存在"
		call endinfo(2)
	else		
%>
<table  width="420" height=300 border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR>
	<TD BACKGROUND="images/topbg.gif" height=9 colspan="2"></TD>
	</TR>
	<tr>
		<td colspan="2" valign=center align=middle height=23 class=big background="images/header.gif">编辑 股民<font color=red><%=rs("帐号")%></font>的资料</td>
	</tr>
	<form method=post  action="?action=SaveEdit">
	<tr>
		<td BGCOLOR="#E4E4EF" width="40%">UID</td>
		<td bgcolor="#FFFFFF" width="60%">&nbsp;<%=rs("id")%><input type=hidden name=Uid value="<%=rs("id")%>"></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">帐户</td>
		<td bgcolor="#FFFFFF">&nbsp;<font color=blue><%=rs("帐号")%></font><input type=hidden name=username value="<%=rs("帐号")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">开户日期</td>
		<td bgcolor="#FFFFFF">&nbsp;<%=rs("开户日期")%></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">最后日期</td>
		<td bgcolor="#FFFFFF">&nbsp;<%=rs("最后日期")%></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">状态</td>
		<td bgcolor="#FFFFFF"><input type=radio name=Locked value="0" <%if rs("锁定")<>1 then%> checked <%end if%>> 正常<input type=radio name=Locked value="1" <%if rs("锁定")=1 then%> checked <%end if%>> 锁定 </td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">资金</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=Crash value="<%=rs("资金")%>"></td>
	</tr>			
	<tr>
		<td BGCOLOR="#E4E4EF">总资金</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TotalFund value="<%=rs("总资金")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><font color="#333333">持股种类</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=StockNum value="<%=rs("持股种类")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>今日买入</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TodayBuy readonly value="<%=rs("今日买入")%>"></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>今日卖出</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TodaySale readonly value="<%=rs("今日买入")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>总买入</font></td> 
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TotalBuy readonly value="<%=rs("总买入")%>"></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>总卖出</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TotalSale readonly value="<%=rs("总卖出")%>"></td>
	</tr>
	<tr><td colspan="2" bgcolor="#FFFFFF" align="center"><input type=submit name=submit value=执行修改></td>
	</form>
</table>	
<%		
	end if
end sub
'---------------保存修改---------
sub SaveEdit()
	dim username,Uid
	username=trim(replace(request.Form("username"),"'",""))
	Uid=request.Form("Uid")
	if username="" or (not isnumeric(uid)) then
		errmess="<li>错误的参数，您所进行的操作非法"
		call endinfo(2)
		exit sub
	end if	
	sql= "select * from 客户 where id="&Uid&""         
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs("资金")=request.form("Crash")
	rs("总资金")=request.form("TotalFund")
	rs("锁定")=request.form("Locked")
	rs("持股种类")=request.form("StockNum")
	rs("今日买入")=request.form("TodayBuy")
	rs("今日卖出")=request.form("TodaySale")
	rs("总买入")=request.form("TotalBuy")
	rs("总卖出")=request.form("TotalSale")		
	rs.update
	rs.close
	sucmess="<li>用户 <font color=blue>"&username&"</font> 的资料更新成功"
	rUrl="?"
	call endinfo(2)
end sub

sub DelUser()
		
end sub
%>