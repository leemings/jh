<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="数据库管理"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
if not master then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
else
	call AdminHead()%>
	<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
	<%select case request("action")
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
	call footer()%>
	</table>
<%end if
'=====================================
sub main()%>
  <tr> 
    <th align="center" colspan="2" height="25">数据库管理</th>
  </tr>
  <tr> 
    <td class=tablebody2 colspan="2" height="25" align="left"><b>基本数据管理</b></td>
	</tr>
	<form method="POST" action="?action=Update1" name="f1">	
	<tr>
    <td class=tablebody1 width="50%"><b>重新计算股票经营权</b><br>根据股民持股数重新计算股票的经营者<br>建议每隔一段时间运行一次</td>
    <td class=tablebody1 width="50%"> <input type=submit name="submit" value="更新经营权"></td>
	</tr>
	</form>
	<form method="POST" action="?action=Update2" name="f2">
	<tr>
    <td class=tablebody1 width="50%"><b>重新计算资产排名</b><br>重新计算用户持股数量、持股现值、总资产值<br>建议每隔一段时间运行一次</td>
    <td class=tablebody1 width="50%"> <input type=submit name="submit" value="刷新资产排名"></td>
	</tr>
	</form>
	<form method="POST" action="?action=Update3" name="f3">
	<tr>
    <td class=tablebody1 width="50%"><b>更新股票数据</b><br></td>
    <td class=tablebody1 width="50%"> <input type=submit name="submit" value="更新股票数据"></td>
	</tr>
	</form>	
	<tr>
    <td class=tablebody2 colspan="2" height="25" align="left"><B>随机事件管理</b></td>
	</tr>	
	<form method="POST" action="?action=ClearEvent" name="f4">
	<tr>
    <td class=tablebody1 width="50%"><b>清空随机事件记录</b><br>删除所有的随机事件记录</td>
    <td class=tablebody1 width="50%"> <input type=submit name="submit" value="执行"></td>
	</tr>
	</form>
	<form method="POST" action="?action=ClearEvent&reaction=created" name="f4">
	<tr>
		<td class="tablebody1" width="50%"><b>重建随机事件表</b><br>删除所有的随机事件记录，重新建立随机事件表[RndEvent]</td> 
		<td class="tablebody1" width="50%"> <input type=submit name="submit" value="执行"></td>
	</tr>
	</form>		
	<tr>
		<td class=tablebody2 colspan="2" height="25" align="left"><b>SOL语句执行操作</b></td>
	</tr>	
	<form method="POST" action="?action=ExecuteSql" name="f3">
	<tr>
		<td class=tablebody1 colspan="2" align=center><br>&nbsp;&nbsp;&nbsp;&nbsp;本操作仅限高级、对SQL编程比较熟悉的用户，您可以直接输入sql执行语句，比如delete from [KeHu] where username='test'进行删除用户，在操作前请慎重考虑您的执行语句是否正确和完整，本操作直接对股票数据库进行操作，执行后不可恢复。<br><br>请输入SQL语句：<input type=text name="sql" size=100> <input type=submit name=submit value="执行"><br><br></td>
	</tr>
	</form>		
<%end sub
'----------重新计算股票经营权-------------
sub Update1()
	dim rst
    set rs=gp_conn.execute("select sid from [GuPiao] order by sid")
	do while not rs.eof
		set rst=gp_conn.execute("select top 1 ZhangHao,Uid from [DaHu] where sid="&rs(0)&" and Uid>0 order by ChiGuShu desc")
		if rst.eof then	
			gp_conn.execute("update [GuPiao] set JingYingZhe=null,Uid=0 where sid="&rs(0))
		else
			gp_conn.execute("update [GuPiao] set JingYingZhe='"&rst(0)&"',Uid="&rst(1)&" where sid="&rs(0))
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
	set rs=gp_conn.execute("select id from [KeHu] where suoding<2 order by id")
	do while not rs.eof 
		TotalFund=0 :	StockTypeNum=0	
		sql="select d.ChiGuShu,g.DangQianJiaGe from [DaHu] d inner join [GuPiao] g on d.sid=g.sid where d.uid="&rs(0)
		set rst=gp_conn.execute(sql)
		do while not rst.eof
			TotalFund=TotalFund+rst(0)*rst(1)
			StockTypeNum=StockTypeNum+1
			rst.movenext
		loop
		gp_conn.execute("update [KeHu] set ZongZiJin=ZiJin+"&TotalFund&",ChiGuZhongLei="&StockTypeNum&" where id="&rs(0))
		rs.movenext
	loop
	rs.close
	sucmess="<li>用户持股数量、持股现值、总资产值更新成功"
	call endinfo(1)
end sub

sub Update3()
	dim StockNum,TongJi
	set rs=gp_conn.execute("select sid,TongJi from [GuPiao] order by sid")
	do while not rs.eof 
		TongJi=rs(1)
		StockNum=gp_conn.execute("select sum(ChiGuShu) from [DaHu] where sid="&rs(0))(0)
		if StockNum="" or isnull(StockNum) then StockNum=0
		if TongJi="" or isnull(TongJi) then TongJi="0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
		gp_conn.execute("update [GuPiao] set ShengYuGuFen=ZongGuFen-"&StockNum&",TongJi='"&TongJi&"' where sid="&rs(0))
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
	gp_conn.Execute(sql)
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
		gp_conn.execute("drop table [RndEvent]")	
		'重建事件表 RndEvent
		sql="CREATE TABLE [RndEvent] (ID int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,content varchar(255),AddTime datetime default now())"
		gp_conn.execute(sql)
		sucmess=sucmess+"<br>"+"<li>重建事件表成功"
	else
		'删除所有随机事件
		gp_conn.execute("delete from [RndEvent]")
		sucmess=sucmess+"<br>"+"<li>所有随机事件已经清空" 
	end if
	call endinfo(1)	
end sub
%>