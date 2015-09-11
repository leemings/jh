<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head><title><%=Gupiao_Setting(5)%>-股市控制面板</title>
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
		case "Update1"		
			call Update1()
		case "Update2"		
			call Update2()
		case "Update3"		
			call Update3()			
		case "ExecuteSql"		
			call ExecuteSql()
		case "ClearEvent"
			 call ClearEvent()
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
%>
<table width="97%" border="0" cellspacing="1" cellpadding="3" align="center" style="border: 1px #6595D6 solid; background-color: #FFFFFF;">
	<tr>
		<td align="center" colspan="2" height="25" background="Images/Header.gif"><b>数据库管理</b></td>
	</tr>
	<tr>
		<th colspan="2" height="25" class="admin" align="left">基本数据管理</th>
	</tr>
	<form method="POST" action="?action=Update1" name="f1">	
	<tr>
		<td class="forumRow" width="50%"><b>重新计算股票经营权</b><br>根据股民持股数重新计算股票的经营者<br>建议每隔一段时间运行一次</td> 
		<td class="forumRow" width="50%">
		<input type=submit name="submit" value="更新经营权">
		</td>
	</tr>
	</form>
	<form method="POST" action="?action=Update2" name="f2">
	<tr>
		<td class="forumRow" width="50%"><b>重新计算资产排名</b><br>重新计算用户持股数量、持股现值、总资产值<br>建议每隔一段时间运行一次</td> 
		<td class="forumRow" width="50%">
		<input type=submit name="submit" value="刷新资产排名">
		</td>
	</tr>
	</form>
	<form method="POST" action="?action=Update3" name="f3">
	<tr>
		<td class="forumRow" width="50%"><b>更新股票数据</b><br></td> 
		<td class="forumRow" width="50%">
		<input type=submit name="submit" value="更新股票数据">
		</td>
	</tr>
	</form>	
	<tr>
		<th colspan="2" height="25" class="admin" align="left">随机事件管理</th>
	</tr>	
	<form method="POST" action="?action=ClearEvent" name="f4">
	<tr>
		<td class="forumRow" width="50%"><b>清空随机事件记录</b><br>删除所有的随机事件记录</td> 
		<td class="forumRow" width="50%">
		<input type=submit name="submit" value="执行">
		</td>
	</tr>
	</form>
	<form method="POST" action="?action=ClearEvent&reaction=created" name="f4">
	<tr>
		<td class="forumRow" width="50%"><b>重建随机事件表</b><br>删除所有的随机事件记录，重新建立随机事件表[RndEvent]</td> 
		<td class="forumRow" width="50%">
		<input type=submit name="submit" value="执行">
		</td>
	</tr>
	</form>		
	<tr>
		<th colspan="2" height="25" class="admin" align="left">SOL语句执行操作</th>
	</tr>	
	<form method="POST" action="?action=ExecuteSql" name="f3">
	<tr>
		<td class="forumRow" colspan="2">
		<br>&nbsp;&nbsp;&nbsp;&nbsp;本操作仅限高级、对SQL编程比较熟悉的用户，您可以直接输入sql执行语句，比如delete from [客户] where username='test'进行删除用户，在操作前请慎重考虑您的执行语句是否正确和完整，本操作直接对股票数据库进行操作，执行后不可恢复。<br><br>
 		请输入SQL语句：
		<input type=text name="sql" size=100>
		<input type=submit name=submit value="执行">
		<br><br>		
		</td>
	</tr>
	</form>		
</table>	
<% 
end sub
'----------重新计算股票经营权-------------
sub Update1()
	dim rst
    set rs=conn.execute("select sid from [股票] order by sid")
	do while not rs.eof
		set rst=conn.execute("select top 1 帐号,Uid from [大户] where sid="&rs(0)&" and Uid>0 order by 持股数 desc")
		if rst.eof then	
			conn.execute("update [股票] set 经营者=null,Uid=0 where sid="&rs(0))
		else
			conn.execute("update [股票] set 经营者='"&rst(0)&"',Uid="&rst(1)&" where sid="&rs(0))
		end if	
		rs.movenext
	loop
	rst.close
	rs.close	
	sucmess="<li>数据库更新成功"
	call endinfo(1)
end sub
'-----------重新计算资产排名-------------
sub Update2()
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
	rst.close
	rs.close
	sucmess="<li>用户持股数量、持股现值、总资产值更新成功"
	call endinfo(1)
end sub

sub Update3()
	dim StockNum,TongJi
	set rs=conn.execute("select sid,TongJi from [股票] order by sid")
	do while not rs.eof 
		TongJi=rs(1)
		StockNum=conn.execute("select sum(持股数) from [大户] where sid="&rs(0))(0)
		if StockNum="" or isnull(StockNum) then StockNum=0
		if TongJi="" or isnull(TongJi) then TongJi="0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
		conn.execute("update [股票] set 剩余股份=总股份-"&StockNum&",TongJi='"&TongJi&"' where sid="&rs(0))
		rs.movenext
	loop
	rs.close
	sucmess="<li>更新成功"
	call endinfo(1)
end sub
'------SQL语句执行操作-----------
sub ExecuteSql()
	if request("sql")="" then
		errmess=errmess+"<br>"+"<li>请输入SQL语句" 
		call endinfo(1)		
		exit sub
	else
		sql=request("sql")
	end if	
	
	On Error Resume Next 
	conn.Execute(sql)
	if err.number="0" then
		sucmess=sucmess+"<br>"+"<li>输入的SQL语句："&sql
		sucmess=sucmess+"<br>"+"<li>语句执行成功"
	 	call endinfo(1)
	else
		errmess=errmess+"<br>"+"<li>输入的SQL语句："&sql
		errmess=errmess+"<br>"+"<li>语句有问题，具体出错如下："
		errmess=errmess+"<br><li>"&Err.Description
		err.clear
		call endinfo(1)
	end if		
	
end sub

sub ClearEvent()
	if request("reaction")="created" then
		'删除事件表[事件]
		conn.execute("drop table [RndEvent]")	
		'重建事件表 RndEvent
		sql="CREATE TABLE [RndEvent] (ID int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,content varchar(255),AddTime datetime default now())"
		conn.execute(sql)
		sucmess=sucmess+"<br>"+"<li>重建事件表成功"
	else
		'删除所有随机事件
		conn.execute("delete from [RndEvent]")
		sucmess=sucmess+"<br>"+"<li>所有随机事件已经清空" 
	end if
	call endinfo(1)	
end sub
%>