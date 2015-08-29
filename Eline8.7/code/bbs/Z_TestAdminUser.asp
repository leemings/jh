<!--#include file="conn.asp"-->
<!--#include file="z_testconn.asp"-->

<!-- #include file="inc/const.asp" -->
<%
'------------------------
'script writed by 我来了
'2003-01-06     幻海晴空
'------------------------
	stats="开心辞典 用户管理"
	
	if not master then
		errmsg=errmsg+"<br><li>您没有在开心辞典管理用户的权限，请确认你已经登录并且具有相应的权限！"
		founderr=true
		stats="开心辞典 错误信息"
		call nav()
		call head_var(0,0,"开心辞典","z_test.asp")		
		call dvbbs_error()
	else
		call nav()
		call head_var(0,0,"开心辞典","z_test.asp")
		dim ars
		select case request("action")
			case "EditUser"
				call EditUser()
			case "SaveEdit"
				call SaveEdit()
			case "Clean"
				call Clean()
			case "del"
				call Del()									
			case else		
				call UserList()
		end select
	end if
	set aconn=nothing
	call activeonline()
	call footer()
'================================================
'-----------------开心辞典用户列表--------------
sub UserList()
%>
<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
	<tr><th>欢迎 <%=membername%> 进入开心辞典 管理中心</th></tr>
	<tr><td class=tablebody2> 
      <%
	call Admin_Head()
	
	dim orders,paixu
	if not isInteger(request("orders")) or request("orders")="" then
		orders=0
	else
		orders=cint(request("orders"))
	end if	
	select case orders
		case 1
			sql=" order by userok"
		case 2
			sql=" order by userans"
		case 3
			sql=" order by userup"
		case 4
			sql=" order by userpass"
		case 5
			sql=" order by usersinew"
		case 6
			sql=" order by intime"
		case 7
			sql=" order by username"			
		case else
			sql=" order by id"			
	end select
	if request("paixu")="" or (not isnumeric(request("paixu"))) then
		paixu=1
	else
		paixu=request("paixu")
	end if
	if cint(paixu)=1 then sql=sql&" desc"	
%>
	
	<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
		<tr><td align="left" class=tablebody2>
		&nbsp;<font color=blue>□</font> 点击用户ID可以进入辞典用户属性编辑&nbsp;&nbsp;<font color=blue>□</font> 点击用户名可以查看该用户的论坛属性 <br>
		&nbsp;<font color=blue>□</font> 可以删除的用户是指从来没有上传过题目或者没有答过题目的用户，这些记录完全是无用的，可以放心清理
		</td></tr>
	</table><table border=0><tr><td height="3"></td></tr></table>
	<table align=center cellspacing=1 cellpadding=3 class=tableborder1 style="width:97%">
		<tr><th colspan="9">用户管理</th></tr> 
		<tr>
			<td class=tablebody2 colspan="9">
				<table border="0" cellspacing=0 cellpadding=0 width="100%"><tr> 
				<form action="?action=<%=request("action")%>&orders=<%=request("orders")%>" method="post">
				<td>排序方式： 
				<select name=paixu onchange='javascript:submit()'>
					<option value=1 <%if cint(paixu)=1 then%> selected <%end if%>>降序</option>
					<option value=0 <%if cint(paixu)=0 then%> selected <%end if%>>升序</option>
				</select>
				</td> 
				</form>
				<form action="?action=EditUser" method="post">
				<td>
					用户名：&nbsp; <input type="text" name="UserName" value="">&nbsp;&nbsp;<input type="submit" value="搜索" name="search">
				</td>
				</form>
				</form>
				<form action="?action=Clean" method="post">
				<td>
					共有<font color=blue><%=aconn.execute("select count(id) from [testuser]")(0)%></font>个用户，其中<font color=red><%=aconn.execute("select count(id) from [testuser] where userup=0 and userans=0")(0)%></font>个可以删除&nbsp;<input type="submit" value="清理" name="submit">
				</td>
				</form>				
				</tr></table>
			</td> 
		</tr>	
        <tr>
			<td class=TopLighNav1 align="center"><%if orders=0 then%><font color=<%=Forum_body(8)%>>ID</font><%else%><a href=?action=userlist&orders=0&paixu=<%=paixu%>>ID</a><%end if%></td> 
			<td class=TopLighNav1 align="center"><%if orders=7 then%><font color=<%=Forum_body(8)%>>用户名</font><%else%><a href=?action=userlist&orders=7&paixu=<%=paixu%>>用户名</a><%end if%></td> 
			<tD class=TopLighNav1 align="center"><%if orders=2 then%><font color=<%=Forum_body(8)%>>答题数</font><%else%><a href=?action=userlist&orders=2&paixu=<%=paixu%>>答题数</a><%end if%></td>  
			<td class=TopLighNav1 align="center"><%if orders=1 then%><font color=<%=Forum_body(8)%>>答对数</font><%else%><a href=?action=userlist&orders=1&paixu=<%=paixu%>>答对数</a><%end if%></td> 
			<td class=TopLighNav1 align="center"><%if orders=3 then%><font color=<%=Forum_body(8)%>>上传数</font><%else%><a href=?action=userlist&orders=3&paixu=<%=paixu%>>上存数</a><%end if%></td> 
			<td class=TopLighNav1 align="center"><%if orders=4 then%><font color=<%=Forum_body(8)%>>通过数</font><%else%><a href=?action=userlist&orders=4&paixu=<%=paixu%>>通过数</a><%end if%></td>  
			<td class=TopLighNav1 align="center"><%if orders=6 then%><font color=<%=Forum_body(8)%>>最后时间</font><%else%><a href=?action=userlist&orders=6&paixu=<%=paixu%>>最后时间</a><%end if%></td>
			<td class=TopLighNav1 align="center"><%if orders=5 then%><font color=<%=Forum_body(8)%>>游戏币</font><%else%><a href=?action=userlist&orders=5&paixu=<%=paixu%>>游戏币</a><%end if%></td>
			<td class=TopLighNav1 align="center">操作</td>
        </tr>		
        <%
		dim currentpage,page_count,Pcount
		dim totalrec,endpage
		currentPage=request("page")
		if currentpage="" or not isInteger(currentpage) then
			currentpage=1
		else
			currentpage=clng(currentpage)
			if err then
				currentpage=1
				err.clear
			end if
		end if
		Set ars= Server.CreateObject("ADODB.Recordset")
		sql="select * from [testuser] "& sql
		ars.open sql,aconn,1,1
		if ars.eof and ars.bof then
			response.write "<tr><td colspan=8 class=tablebody1>没有找到相关记录。</td></tr>"
		else
			ars.PageSize = Cint(Forum_Setting(11)) 
			ars.AbsolutePage=currentpage
			page_count=0
			totalrec=ars.recordcount
			do while (not ars.eof) and (not page_count = Cint(Forum_Setting(11)))
%>
			  <tr>
			  	<td align="center" class=tablebody1><a href="?action=EditUser&id=<%=ars("id")%>"><%=ars("id")%></a></td>
				<td align="center" class=tablebody1><a href=dispuser.asp?name=<%=ars("username")%> target=_blank><%=ars("username")%></a></td>
				<td class=tablebody1 align="center"><%=ars("userans")%></td>
				<td class=tablebody1 align="center"><%=ars("userok")%></td> 
				<td class=tablebody1 align="center"><%=ars("userup")%></td>
				<td class=tablebody1 align="center"><%=ars("userpass")%></td> 
				<td class=tablebody1 align="center"><%=ars("intime")%></td>
				<td class=tablebody1 align="center"><%=ars("usersinew")%></td>
				<td class=tablebody1 align="center"><a href="?action=EditUser&id=<%=ars("id")%>">编辑</a> | <a href="?action=del&id=<%=ars("id")%>" onclick="javascript:{if(confirm('您确定删除该用户吗?')){return true;}return false;}">删除</a></td>
			  </tr>
<%			
				page_count = page_count + 1		
				ars.movenext
			loop
			Pcount=ars.PageCount
%>
			<tr><td colspan=9 class=TopLighNav1 align=right>分页：
<%
			if currentpage > 4 then
				response.write "<a href=""?page=1&action="&request("action")&"&orders="&request("orders")&"&paixu="&paixu&""">[1]</a> ..."
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
				response.write " <a href=""?page="&i&"&action="&request("action")&"&orders="&request("orders")&"&paixu="&paixu&""">["&i&"]</a>"
				end if
			end if
			next
			if currentpage+3 < Pcount then 
				response.write "... <a href=""?page="&Pcount&"&action="&request("action")&"&orders="&request("orders")&"&paixu="&paixu&""">["&Pcount&"]</a>"
			end if
				
		end if
		ars.close
		set ars=nothing	
%>
        		  
		</table> 
		<br>
	</td></tr>	
</table>	
<%	
end sub
'--------------------------
sub EditUser()
	dim id,UserName
	id=request("id")
	UserName=replace(trim(request.Form("UserName")),"'","")
	if id="" and request("search")="搜索" then
		if UserName="" then
			errmsg=errmsg+"<br><li>请输入用户名"
			call test_err()
			exit sub
		else
			sql=" UserName='"& UserName &"'"
		end if					
	elseif not isInteger(id) then
		errmsg=errmsg+"<br><li>请指定用户ID"
		call test_err()
		exit sub
	else
		sql=" id="&id&""	
	end if

	sql="select * from [testuser] where "&sql
	set ars=aconn.execute(sql)
	if ars.eof and not ars.bof then
		errmsg=errmsg+"<br><li>没有找到指定的用户"
		founderr=true
		ars.close
		call test_err()
		exit sub
	end if
%>

	<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
	<form method="POST" action="?action=SaveEdit">
	<tr><th>欢迎 <%=membername%> 进入开心辞典 管理中心</th></tr>
	<tr><td class=tablebody2>
        <%call Admin_Head()%>
        
<table align=center cellspacing=1 cellpadding=3 class=tableborder1 style="width:97%">
  <tr><th colspan=8>编辑用户 <%=ars("username")%> 的属性</th></tr> 
  <tr>
	<td width="35%" class=tablebody2>ID：</td>
	<td width="65%" class=tablebody1>&nbsp;<%=ars("id")%><input type="hidden" name="id" value=<%=ars("id")%>></td>
  </tr>
  <tr>	
	<td class=tablebody2>用户名：</td> 
	<td class=tablebody1>&nbsp;<%=ars("username")%></td>
  </tr>
  <tr>	
	<td class=tablebody2>最后答题时间：</td> 
	<td class=tablebody1>&nbsp;<%=ars("intime")%></td>
  </tr>
   <tr>	
	<td class=tablebody2>已记发帖数：</td> 
	<td class=tablebody1>&nbsp;<%=ars("ftl")%> 篇</td>
  </tr> 
  <tr>	
    <td class=tablebody2>游戏币：</td>
    <td class=tablebody1>&nbsp;<input type="text" name="usersinew" value=<%=ars("usersinew")%>> 个</td>
  </tr>  
  <tr>
	<td class=tablebody2>答题总数：</td>
	<td class=tablebody1>&nbsp;<input type="text" name="userans" value=<%=ars("userans")%>> 题</td>
  </tr>
  <tr>	
	<td class=tablebody2>答对数：</td>
	<td class=tablebody1>&nbsp;<input type="text" name="userok" value=<%=ars("userok")%>> 题</td>  
  </tr>
  <tr>	
    <td class=tablebody2>剩余答题数：</td>
    <td class=tablebody1>&nbsp;<input type="text" name="js" value=<%=ars("js")%>> 题</td> 
  </tr>
  <tr>	
    <td class=tablebody2>上传题目：</td>
    <td class=tablebody1>&nbsp;<input type="text" name="userup" value=<%=ars("userup")%>> 题</td>
  </tr>
  <tr>	
    <td class=tablebody2>上传通过：</td>
    <td class=tablebody1>&nbsp;<input type="text" name="userpass" value=<%=ars("userpass")%>> 题</td>
  </tr> 
  <tr>	
    <td align="center" class=tablebody1 colspan="2"><input type="submit" value="提交" name="submit">&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" value="重置" name="reset"></td>
  </tr>        
</table>
</form></table>   
<%
		set ars=nothing 	
end sub

sub SaveEdit()
	dim id
	id=request.Form("id")
	if id="" or (not isnumeric(id)) then
		errmsg=errmsg+"<br><li>请指定用户ID"
		call test_err()
		exit sub
	end if		
	if request("usersinew")="" or (not isnumeric(request("usersinew"))) then
		errmsg=errmsg+"<br><li>游戏币不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("usersinew")<0 then 
		errmsg=errmsg+"<br><li>游戏币不能为负数，请重新输入"
		founderr=true	
	end if
	if request("userans")="" or (not isnumeric(request("userans"))) then
		errmsg=errmsg+"<br><li>答题总数不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("userans")<0 then 
		errmsg=errmsg+"<br><li>答题总数不能为负数，请重新输入"
		founderr=true	
	end if
	if request("userok")="" or (not isnumeric(request("userok"))) then
		errmsg=errmsg+"<br><li>答对数不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("userans")<0 then 
		errmsg=errmsg+"<br><li>答对数不能为负数，请重新输入"
		founderr=true	
	end if
	if request("js")="" or (not isnumeric(request("js"))) then
		errmsg=errmsg+"<br><li>剩余答题数不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("js")<0 then 
		errmsg=errmsg+"<br><li>剩余答题数不能为负数，请重新输入"
		founderr=true	
	end if
	if request("userup")="" or (not isnumeric(request("userup"))) then
		errmsg=errmsg+"<br><li>上传题目数不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("userup")<0 then 
		errmsg=errmsg+"<br><li>上传题目数不能为负数，请重新输入"
		founderr=true	
	end if
	if request("userpass")="" or (not isnumeric(request("userpass"))) then
		errmsg=errmsg+"<br><li>上传通过数不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("userpass")<0 then 
		errmsg=errmsg+"<br><li>上传通过数不能为负数，请重新输入"
		founderr=true	
	end if
	if founderr then 
		call test_err()
		exit sub
	end if
	aconn.execute("update [testuser] set usersinew="&request("usersinew")&",userans="&request("userans")&",userok="&request("userok")&",js="&request("js")&",userup="&request("userup")&",userpass="&request("userpass")&" where id="&id)
	sucmsg=sucmsg+"<br><li>操作成功"
	call suc()
end sub

sub Clean()
	aconn.execute("delete from [testuser] where userup=0 and userans=0")
	sucmsg=sucmsg+"<br><li>操作成功，已将所有无用记录清除"
	call suc()	
end sub

sub Del()
	dim id
	id=request.QueryString("id")
	if id="" or (not isnumeric(id)) then
		errmsg=errmsg+"<br><li>请指定用户ID"
		call test_err()
		exit sub
	end if
	aconn.execute("delete from [testuser] where id="&id&"")
	sucmsg=sucmsg+"<br><li>操作成功，一个用户已删除"
	call suc()
end sub
%>