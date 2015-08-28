<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head><title><%=Gupiao_Setting(5)%>-客户管理</title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
	
<% 
	if request("action")="Reflash" then
		call Reflash()
	else	
		call main()
	end if	
response.Write "</body></html>"
CloseDatabase		'关闭数据库  
'=====================================
sub main() 
	dim Title 
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
		Title="查找结果"
	else
		if request("paixu")="资产" then
			sql= "select * from 客户 order by 资金 desc"
			Title="资产排行"	
		else 
			Title="总资产排行"
			sql= "select * from 客户 order by 总资金 desc"
		end if	
	end if					
%>
<table  width="98%" border=0 cellspacing=1 cellpadding=0 align=center bgcolor="#0066CC">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
	</TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif"><font size="2"><b>股市大鳄排行榜</b></font></td>
	</tr>
	<tr><td bgcolor="#E4E8EF"> 
<br>
<table cellspacing=1 cellpadding=3 align=center width="97%" bgcolor="#0066CC">
	<tr>
		<td align=left valign=middle background="Images/title.gif" height="21"> &nbsp;<a href="stock.asp">股票交易大厅</a> | <a href="?paixu=总资产">总资金排行</a> | <a href="?paixu=资产">资金排行</a> | <a href=javascript:history.go(-1)>返回上一页</a></td>
	</tr> 
</table>
<br>

<table border="0" width="97%" bgcolor="#0066CC" cellspacing="1" cellpadding="3" align="center">
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif" colspan="11"><b><%=Title%></b></td>
	</tr>
	<tr bgcolor="#ffffff"> 
	<%if request("action")="search" then%>
	<td colspan="11" height=18><font color=red>搜索结果如下：</font><a href="?paixu=<%=request("paixu")%>">[返回排行榜]</a></td>  
	<%else%>	
	<form name="form1" method="post" action="?action=search">
		<td colspan="11" height=18>快速查找：<input type="text"  name=username> <input type="checkbox"  name="usernamechk" value="yes" checked>用户名完全匹配<input type="submit" name="Submit" value="搜索"></td>
	</form>
	<%end if%>
	</tr>
	<tr bgcolor="#666666" align="center"> 
		<td background="Images/lan.gif"><font color=#ffffff>名次</font></td>
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
			rs.PageSize = Gupiao_Setting(2) 
			rs.AbsolutePage=currentpage
			page_count=0
			totalrec=rs.recordcount
			do while (not rs.eof) and (not page_count = 20) 
%>
			<tr align=center bgcolor='#E4E6EF'>  
				<td><%if request("action")="search" then%>-<%else%><font color=red><%=Gupiao_Setting(2)*(currentpage-1)+page_count+1%></font><%end if%></td>
				<td><a href="dispu.asp?uid=<%=rs("id")%>&username=<%=rs("帐号")%>"><%=rs("帐号")%></a></td>
				<td align="right"><%=formatnumber(rs("资金"),0,true)%></td>
				<td align="right"><%=formatnumber(rs("总资金"),0,true)%></td>
				<td><%=rs("持股种类")%></td>
				<td><%=rs("今日买入")%></td>
				<td><%=rs("今日卖出")%></td>
				<td><%=rs("总买入")%></td>
				<td><%=rs("总卖出")%></td>
				<td><%=formatdatetime(rs("最后日期"),2)%></td>				
				<td><%if rs("锁定")=1 then%><font color=red>锁定<%else%>正常<%end if%></td>
			</tr>
<%			
				page_count = page_count + 1		
				rs.movenext
			loop
			Pcount=rs.PageCount
%>
			<tr bgcolor="#666666" align="center"> 
				<td background="Images/lan.gif" colspan=2><font color=#ffffff>单位</font></td>
				<td background="Images/lan.gif" align="right"><font color=#ffffff>元</font></td>
				<td background="Images/lan.gif" align="right"><font color=#ffffff>元</font></td>
				<td background="Images/lan.gif"><font color=#ffffff>种</font></td>
				<td background="Images/lan.gif"><font color=#ffffff>笔</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>笔</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>笔</font></td>
				<td background="Images/lan.gif"><font color=#ffffff>笔</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>年-月-日</font></td> 		
				<td background="Images/lan.gif"><font color=#ffffff>-</font></td> 
			</tr>
			<tr><td colspan=11 background="images/title.gif" align=right>
			<table border="0" cellpadding="0" cellspacing="0" width="100%"><tr>
			<td align=left>共有<font color=blue><%=totalrec%></font>个股民，每页<font color=blue><%=Gupiao_Setting(2)%></font>个，第<font color=red><%=currentpage%></font>页/共<font color=blue><%=Pcount%></font>页</td>
			<td align=right>分页：
<%
			if currentpage > 4 then
				response.write "<a href=""?page=1&paixu="&request("paixu")&""">[1]</a> ..."
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
				response.write " <a href=""?page="&i&"&paixu="&request("paixu")&""">["&i&"]</a>"
				end if
			end if
			next
			if currentpage+3 < Pcount then 
				response.write "... <a href=""?page="&Pcount&"&paixu="&request("paixu")&""">["&Pcount&"]</a>"
			end if
			response.Write "</td></tr></table></td></tr>"		
		end if
		rs.close
		set rs=nothing	
%>
 </table>
 <br>
</td></tr>
<tr><td height=32 background="images/footer.gif" align=center><a href="stock.asp">返回股市</a></td></tr>
</table>
<% 
end sub

sub Reflash()
	if session("GP_ReflashFlag")<>"" then
		errmess="<li>现在的数据已经是最新的了，请不要多次刷新"
		call endinfo(1) 
	else
		dim rst,TotalFund,StockTypeNum
		set rs=conn.execute("select id from [客户] order by id")
		do while not rs.eof 
			TotalFund=0 :	StockTypeNum=0	
			sql="select d.持股数,g.当前价格 from [大户] d inner join [股票] g on d.sid=g.sid where d.uid="&rs(0)
			set rst=conn.execute(sql)
			do while not rst.eof
				TotalFund=TotalFund+rst(0)*rst(1)
				StockTypeNum=StockTypeNum+1
				rst.movenext
			loop
			conn.execute("update [客户] set 总资金=资金+"&TotalFund&",持股种类="&StockTypeNum&" where id="&rs(0))
			rs.movenext
		loop
		session("GP_ReflashFlag")="haveFlash"
		response.Redirect Request.ServerVariables("HTTP_REFERER")
	end if		
end sub
%>