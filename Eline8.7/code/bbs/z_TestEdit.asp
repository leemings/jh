<!--#include file="conn.asp"-->
<!--#include file="Z_TestConn.ASP"-->
<!-- #include file="inc/const.asp" -->
<%
REM 开心辞典 已有题库维护
REM 我来了 修改为final版 
REM 原作者不清楚

dim ars 

if not master then  
	errmsg=errmsg+"<br><li>您没有在开心辞典维护已有题库的权限，请确认你已经登录并且具有相应的权限！" 
	founderr=true
	stats="开心辞典 错误信息"
	call nav()
	call head_var(2,0,"","")		
	call dvbbs_error()
else
	stats="开心辞典 已有题库管理"
	call nav()
	call head_var(0,0,"开心辞典","Z_test.asp")
%>
	<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
	<tr><th>欢迎 <%=membername%> 进入开心辞典 管理中心</th></tr>
	<tr><td class=tablebody2>
<%
	call Admin_Head()
	select case request("action")
		case "saveedit"
			call saveedit()
		case "edit"
			call edit()
		case "del"
			call del()
		case "selaction"
			call selaction()	'批量操作	
		case else
			call main()
	end select
	response.Write "</td></tr></table>"
	set aconn=nothing	
end if
call activeonline()
call footer()	
'===================================	
sub main()						
%>
<table align=center cellspacing=1 cellpadding=3 class=tableborder1 style="width:97%">
	<tr><th colspan="6">已有题库管理</th></tr> 
	<tr><td colspan="6" class=TopLighNav1><font color=<%=Forum_body(8)%>>快速搜索</font></td></tr> 
	<form method="POST" action="?action=edit" name=form1>
	<tr>
		<td class=tablebody1 colspan="2">&nbsp;题号: </td>
		<td class=tablebody1 colspan="4"><input type="text" name="Tid" size="20">&nbsp;&nbsp;<input type=submit value="查找" name="B1"></td>
	</tr>
	</form>
	<tr><td colspan=6 class=tablebody2 height=5></td></tr>
<%if request("action")="" or request("userSearch")="" then%>	
	<tr><td colspan="6" class=TopLighNav1><font color=<%=Forum_body(8)%>>高级查询</font></td></tr> 
	<form method="POST" action="?action=userSearch" name=form2>
	<tr>
		<td class=tablebody1 width=20%  colspan="2">&nbsp;上传者</td> 
		<td class=tablebody1 width=80%  colspan="4"><input type="text" name="username" size="20">&nbsp;<input type=checkbox name="usernamechk" value="yes" checked>用户名完整匹配</td>
	</tr>
	<tr>
		<td class=tablebody1 width=20%  colspan="2">&nbsp;题目类型</td> 
		<td class=tablebody1 width=80%  colspan="4">
			<select size=1 name="TitleClass">
				<option value=0>任意</option> 
				<option value=1>选择题</option>
				<option value=2>填空题</option>
			</select>	
		</td>
	</tr>
	<tr>
		<td class=tablebody1 width=20%  colspan="2">&nbsp;题目类别</td> 
		<td class=tablebody1 width=80%  colspan="4">
			<select size=1 name="TitleLb">
				<option value=0>任意</option> 
<%
				set ars=aconn.execute("select id,lb from testlb order by id")
				do while not ars.eof
					response.write "<option value="&ars(0)&">"&ars(1)&"</option>"
					ars.movenext
				loop
				ars.close
				set ars=nothing
%>
			</select>	
		</td>
	</tr>
	<tr>
		<td class=tablebody1 width=20%  colspan="2">&nbsp;题目中包含</td> 
		<td class=tablebody1 width=80%  colspan="4"><input type=text name="Title" size=45></td>
	</tr>
	<tr>
		<td class=tablebody1 width=20%  colspan="2">&nbsp;排序方式</td> 
		<td class=tablebody1 width=80%  colspan="4">
			<select size=1 name="ordername">
				<option value="id">按题号</option> 
				<option value="title">按题目</option>
				<option value="lb">按类别</option>
				<option value="username">按上传者</option>
				<option value="tcount">按应答人数</option>
				<option value="okcount">按答对人数</option>
				<option value="jj">按奖金的多少</option>  
				<option value="time">按时间限制的多少</option> 
			</select>
			<select size=1 name="order">
				<option value="esc">升序</option> 
				<option value="desc">降序</option>
			</select>					
		</td>
	</tr>
	<tr>
		<td class=tablebody2 align=center colspan=6><input name="submit" type=submit value="   搜  索   "></td>
	</tr>
	<input type=hidden name="userSearch" value="2">
	</form>
	
	<tr><td colspan=6 class=tablebody1 height=5></td></tr>
	<tr><td colspan="6" class=TopLighNav1><font color=<%=Forum_body(8)%>>题库自动清理</font>→查找并删除重复的题目</td></tr>  
	<form method="POST" action="?action=userSearch" name=form3>
	<tr>
		<td class=tablebody1 colspan="2">&nbsp;操作说明: </td>
		<td class=tablebody1 colspan="4">
			&nbsp;⑴ 本操作将会把重复的题目删除，只留下其中一题，删除的题目由下面的删除选项决定<br>
			&nbsp;⑵ 本操作只能找出题目完全一致的重复题目<br>
		 	&nbsp;⑶ 如果题目中有单引号，程序将会跳过此题目，所以那些含有单引号的重复题目只能手动删除<br>
			&nbsp;⑷ <font color=<%=Forum_body(8)%>>本操作将会严重占用服务器资源，CPU占100%，所以不要在服务器上直接进行本操作</font><br>
			&nbsp;⑸ 本操作可能会花费比较长的时间，具体要看题库的大小和作为服务器的机器的性能<br>   
			&nbsp;⑹ <font color=blue>对于重复题目不多的题库，建议在<u>高级查询</u>中选择按<u>题目</u>排序找出重复的题目删除</font>
		</td> 
	</tr> 
	<tr>
		<td class=tablebody1 colspan="2">&nbsp;删除选项: </td>
		<td class=tablebody1 colspan="4"><input type=radio name="DeleteSel" value="1" checked>删除题目ID大的题目	<input type=radio name="DeleteSel" value="2">删除题目ID小的题目</td>
	</tr> 
	<tr>
		<td class=tablebody2 align=center colspan=6><input name="submit" type=submit value="开始清理"></td>
	</tr>	
	<input type=hidden name="userSearch" value="1">
	</form>	
<%
elseif request("userSearch")=1 then
	call ClearTitle()	'清理重复的题 
elseif request("userSearch")=2 then
%>
	<tr>
		<td colspan=6 class=TopLighNav1 height=23><font color=<%=Forum_body(8)%>>搜索结果</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="?" class="cblue">高级查询</a></td>  
	</tr>
	<tr><td colspan=6 class=tablebody2 height=5></td></tr>	
<%
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	dim sqlstr
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
	
	Set rs= Server.CreateObject("ADODB.Recordset")
	sqlstr=""
	if trim(request("username"))<>"" then
		if request("usernamechk")="yes" then
			sqlstr=" username='"&trim(request("username"))&"'"
		else
			sqlstr=" username like '%"&trim(request("username"))&"%'"
		end if
	end if
	if cint(request("TitleClass"))>0 then
		if sqlstr="" then
			sqlstr=" tx="&request("TitleClass")-1&""
		else
			sqlstr=sqlstr & " and tx="&request("TitleClass")-1&""
		end if
	end if
	if cint(request("TitleLb"))>0 then
		if sqlstr="" then
			sqlstr=" lb="&request("TitleLb")&""
		else
			sqlstr=sqlstr & " and lb="&request("TitleLb")&""
		end if
	end if	
	if request("Title")<>"" then
		if sqlstr="" then
			sqlstr=" Title like '%"&request("Title")&"%'"
		else
			sqlstr=sqlstr & " and Title like '%"&request("Title")&"%'"
		end if
	end if
	if sqlstr<>"" then sqlstr=" where "&sqlstr
	sqlstr=sqlstr & " order by "&request("ordername")
	if request("order")<>"esc" then sqlstr=sqlstr&" "&request("order")
	
	sql="select * from [test] "& sqlstr
	'response.write sql
	rs.open sql,aconn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=6 class=tablebody1>没有找到相关记录。</td></tr>"
	else		
%>		
	<FORM METHOD=POST ACTION="?action=selaction">
		  <tr>
			<td width="40" align="center" class=TopLighNav1><b>题　号</b></td>
			<td width="45" align="center" class=TopLighNav1><b>题　型</b></td
			><td width="*" align="center" class=TopLighNav1><b>题 目(点击题目进行本题编辑)</b></td>
			<td width="90" align="center" class=TopLighNav1><b>题目类别</b></td>
			<td width="80" align="center" class=TopLighNav1><b>本题上传者</b></td> 
			<td width="110" align="center" class=TopLighNav1><b>操　作</b></td>   
		  </tr>
<%
			rs.PageSize = Cint(Forum_Setting(11)) 
			rs.AbsolutePage=currentpage
			page_count=0
			totalrec=rs.recordcount
			while (not rs.eof) and (not page_count = Cint(Forum_Setting(11)))
%>
			  <tr>
				<td align="center" class=tablebody1><%=rs("id")%></td>
				<td class=tablebody1 align="center"><%if rs("tx")=1 then%>填空题<%else%>选择题<%end if%></td>
				<td class=tablebody1><a href="?action=edit&Tid=<%=rs("id")%>"><%=rs("title")%></a></td> 
				<td class=tablebody1 align="center"><%=GetLbName(rs("lb"))%></td>
				<td class=tablebody1 align="center"><%=rs("username")%></td> 
				<td class=tablebody1 align="center"><a href="?action=edit&Tid=<%=rs("id")%>" title="编辑题目">编辑</a> | <a href="?action=del&Tid=<%=rs("id")%>" title="直接删除本题" onclick="javascript:{if(confirm('您确定删除本题吗?')){return true;}return false;}">删除</a> | <input type="checkbox" name="TitleID" value="<%=rs("id")%>"></td>
			  </tr>
<%
			page_count = page_count + 1
			rs.movenext
			wend
			Pcount=rs.PageCount
%>
			<tr><td colspan=6 class=TopLighNav1 align=center>分页：
<%
			if currentpage > 4 then
				response.write "<a href=""?page=1&userSearch="&request("userSearch")&"&username="&request("username")&"&TitleClass="&request("TitleClass")&"&TitleLb="&request("TitleLb")&"&Title="&request("Title")&"&usernamechk="&request("usernamechk")&"&ordername="&request("ordername")&"&order="&request("order")&"&action="&request("action")&""">[1]</a> ..."
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
				response.write " <a href=""?page="&i&"&userSearch="&request("userSearch")&"&username="&request("username")&"&TitleClass="&request("TitleClass")&"&TitleLb="&request("TitleLb")&"&Title="&request("Title")&"&usernamechk="&request("usernamechk")&"&ordername="&request("ordername")&"&order="&request("order")&"&action="&request("action")&""">["&i&"]</a>"
				end if
			end if
			next
			if currentpage+3 < Pcount then 
				response.write "... <a href=""?page="&Pcount&"&userSearch="&request("userSearch")&"&username="&request("username")&"&TitleClass="&request("TitleClass")&"&TitleLb="&request("TitleLb")&"&Title="&request("Title")&"&usernamechk="&request("usernamechk")&"&ordername="&request("ordername")&"&order="&request("order")&"&action="&request("action")&""">["&Pcount&"]</a>"
			end if
%>		  
<tr><td colspan=5 class=tablebody1 align=center><B>请选择您需要进行的操作</B>：<input type="radio" name="selaction" value=1 checked>批量删除&nbsp;&nbsp;<input type="radio" name="selaction" value=2>移动到类别&nbsp;&nbsp;
<select size=1 name="lbid">
<%
		set ars=aconn.execute("select lb,id from [testlb]")
		do while not ars.eof
			response.write "<option value="&ars(1)&">"&ars(0)&"</option>"
			ars.movenext
		loop
		ars.close
%>
	</select>
	</td>
	<td class=tablebody1 align=center><input type=checkbox value="on" name="chkall" onclick="CheckAll(this.form)">
	</td>
	</tr>
	<tr><td colspan=6 class=tablebody1 align=center>
	<input type=submit name=submit value="执行选定的操作"  onclick="{if(confirm('确定执行选择的操作吗?')){this.document.recycle.submit();return true;}return false;}">
	</td></tr>
	</FORM>		  
<%
	end if 
end if			
%>  					
</table>	
<%
end sub

sub edit()
	dim id,ausername
	dim alb,akey1,akey2,akey3,akey4,ajj,atime,atitle,aok,atx

	id=ltrim(trim(request("Tid")))
	if id="" or (not isnumeric(id)) then
		errmsg=errmsg+"<br><li>请指定正确的参数"
		call test_err()
		exit sub
	end if 

	set ars=server.createobject("adodb.recordset")
	sql="select lb,title,key1,key2,key3,key4,ok,username,id,tx,jj,time from [test] where id="&id
	ars.open sql,aconn,1,1
	if ars.eof and  ars.bof then
		ars.close
		errmsg=errmsg+"<br><li>错误的题目参数，没有找到该题目"
		call test_err()
		exit sub
	end if	
	alb=ars(0)
	atitle=ars(1)
	akey1=ars(2)
	akey2=ars(3)
	akey3=ars(4)
	akey4=ars(5)
	aok=ars(6)
	ausername=ars(7)
	id=ars(8)
	atx=ars(9)
	ajj=ars(10)
	atime=ars(11)
	ars.close
%>
<form method="POST" action='Z_TestEdit.asp?action=saveedit&Tid=<%=id%>'>
<table align=center cellspacing=1 cellpadding="3" class=tableborder1 style="width:97%">
  <tr><td colspan=8 class=TopLighNav1 height=23>编辑题目&nbsp; &nbsp; &nbsp; &nbsp;<a href="Z_TestEdit.asp?action=edit&Tid=<%=id-1%>"><font color=<%=Forum_body(8)%>>上一题</font></a> | <a href="Z_TestEdit.asp?action=edit&Tid=<%=id+1%>"><font color=<%=Forum_body(8)%>>下一题</font></a> | <a href="?">高级查询</a> | <a href="<%=Request.ServerVariables("HTTP_REFERER")%>">返 回</a></td></tr> 
  <tr>
    <td width="70" align="center" class=tablebody2><b>题　号</b></td>
    <td width="82" class=tablebody1>&nbsp;<%=id%> </td>
    <td width="98" class=tablebody2><b>&nbsp;题目所属类别</b></td>
    <td width="102" class=tablebody1><select name="lb"><%
	set ars=server.createobject("adodb.recordset")
	sql="select lb,id from [testlb]"
	ars.open sql,aconn,1,1
	do while not ars.eof
%>
		<option value='<%=ars(1)%>' <%if ars(1)=alb then%>selected<%end if%>><%=ars(0)%></option>
<%
  		ars.movenext
	loop
	ars.close
%></select></td>
    <td width="57" class=tablebody2><b>题型</b></td>
    <td width="63" class=tablebody1><%if atx=1 then%>填空题<%else%>选择题<%end if%></td>
    <td width="90" class=tablebody2><b>本题上传者</b></td>
    <td width="102" class=tablebody1><input type="text" name="username" size="8" value=<%=ausername%>></td>
  </tr>
  <tr >
    <td width="70"  align="center" class=tablebody2><b>题&nbsp;&nbsp;目</b></td>
    <td width="600" colspan="7" align="left" class=tablebody1>&nbsp;<textarea rows="12" name="title" cols="70"><%=atitle%></textarea></td>
  </tr>
  <tr >
    <td width="70"  align="center" class=tablebody2><%if atx=1 then%>正确答案<%else%>答案1<%end if%></td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key1" size="70" value='<%=akey1%>'></td>
  </tr>
  <%if atx=0 then%>
  <tr>
    <td width="70" align="center" class=tablebody2>答案2</td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key2" size="70"  value='<%=akey2%>'></td>
  </tr>
  <tr >
    <td width="70" align="center" class=tablebody2>答案3</td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key3" size="70" value='<%=akey3%>'></td>
  </tr>
  <tr>
    <td width="70"  align="center" class=tablebody2>答案4</td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key4" size="70"  value='<%=akey4%>'></td>
  </tr><%end if%>
  <tr>
    <td width="70"  align="center" class=tablebody2><%if atx=0 then%>正确答案<%end if%></td>
    <td width="81" class=tablebody2>&nbsp;<%if atx=0 then%><input type="text" name="ok" size="10"  value='<%=aok%>'><%end if%></td>
    <td width="116" colspan="2" class=tablebody2 align=center>限时 <input type="text" name="atime" size="7" value=<%=atime%>> 秒</td>
    <td width="172" colspan="2" class=tablebody2 align=center>本题奖金<input type="text" name="jj" size="7" value=<%=ajj%>>元</td>
    <td width="231" colspan="2" class=tablebody2></td>
  </tr>
</table>
<tr><td align=center class=tablebody2><input type="submit" value="修改本题" name="reaction">　　<input type="submit" value="删除本题" name="reaction"></td></tr>
</form>
<%
end sub	

sub saveedit()
	dim id
	id=request("Tid")
	if id="" or (not isnumeric(id)) then
		errmsg=errmsg+"<br><li>请指定正确的题号"
		call test_err()
		exit sub	
	end if
	if request("reaction")="修改本题" then
		dim ausername
		dim lb,key(4),jj,atime,title,ok,tx		
		lb=request("lb"):		jj=request("jj"):		atime=request("atime"):		title=request("title")
		key(0)=request("key1"):		key(1) =request("key2"):		key(2) =request("key3"):		key(3) =request("key4")
		ok=request("ok"):		tx=request("tx"):		ausername=trim(replace(request("username"),"'",""))
			
		if jj="" or (not isnumeric(jj)) then
			errmsg=errmsg+"<br><li>奖金必须填数字"
			founderr=true
		end if
		if atime="" or (not isnumeric(atime)) then
			errmsg=errmsg+"<br><li>时限必须填数字"
			founderr=true
		elseif atime<=0 then
			errmsg=errmsg+"<br><li>时限必须大于0"
			founderr=true			
		end if
		if cint(tx)=0 then
			if ok="" or (not isnumeric(ok)) then
				errmsg=errmsg+"<br><li>请填写正确答案的代号"
				founderr=true
			elseif ok>4 or ok<1 then
				errmsg=errmsg+"<br><li>正确答案的代号只能填写1～4的数字"
				founderr=true			
			end if
		end if	
		if key(0)="" then 
			errmsg=errmsg+"<br><li>第一个备选答案不能为空"
			founderr=true
		end if 							
		if title="" then
			errmsg=errmsg+"<br><li>题目不能为空"
			founderr=true		
		end if
				
		set ars=server.createobject("adodb.recordset")
		sql="select * from [test] where id="&id
		ars.open sql,aconn,1,3
		if ars.eof and ars.bof then
			errmsg=errmsg+"<br><li>错误的题号或者该题目已经删除"
			call test_err()
			exit sub		
		end if
		ars("lb")=lb
		ars("jj")=jj
		ars("time")=atime
		ars("title")=title
		ars("key1")=key(0)
		ars("username")=ausername		
		if ars("tx")=0 then			'对于选择题
			ars("key2")=key(1) 
			ars("key3")=key(2) 
			ars("key4")=key(3) 
			ars("ok")=ok
		end if
		ars.update
		ars.close
		set ars=nothing
		sucmsg=sucmsg+"<br><li>编辑题目成功，请返回进行其他操作"
		'rUrl="Z_TestAdminUpLoad.asp"
		call suc()
	elseif request("reaction")="删除本题" then
	    rUrl="Z_Testedit.asp"
		call del()
	end if
end sub

sub del()
	dim id
	id=request("Tid")
	if id="" or (not isnumeric(id)) then
		errmsg=errmsg+"<br><li>请指定正确的题号"
		call test_err()
		exit sub	
	end if
	aconn.execute("delete from [test] where id="&id)
	sucmsg=sucmsg+"<br><li>删除题目成功，请返回进行其他操作"
	call suc()
end sub

sub selaction()
	if request("selaction")="" then
		errmsg=errmsg+"<br><li>请指定相关参数"
		call test_err()
		exit sub
	end if
	if request("selaction")=1 then		'批量删除
		aconn.execute("delete from [test] where id in ("&replace(request("TitleID"),"'","")&")")
		sucmsg=sucmsg+"<br><li>操作成功，请返回进行其他操作"
		call suc()	
	elseif request("selaction")=2 then '批量移到到类别
		dim lbid
		lbid=request("lbid")
		aconn.execute("update [test] set lb="&lbid&" where id in("&replace(request("TitleID"),"'","")&")")
		sucmsg=sucmsg+"<br><li>操作成功，请返回进行其他操作"
		call suc()	
	else
		errmsg=errmsg+"<br><li>错误的参数"
		call test_err()
		exit sub	
	end if
end sub
'================清理重复的题目====================
sub ClearTitle()
	Server.ScriptTimeOut=9999999
	set ars=server.createobject("adodb.recordset")
	dim startime1
	startime1=timer()
	if cint(request("DeleteSel"))=1 then
		sql="select id,title from [test] order by id"
		ars.open sql,aconn,1,1
		on error resume next
		for i=1 to ars.recordcount
			aconn.execute("delete from [test] where title='"&ars(1)&"' and id>"&ars(0))
			ars.movenext
		next
		ars.close		
	else
		sql="select id,title from [test] order by id desc"
		ars.open sql,aconn,1,1
		on error resume next
		for i=1 to ars.recordcount
			aconn.execute("delete from [test] where title='"&ars(1)&"' and id<"&ars(0))
			ars.movenext
		next
		ars.close		
	end if	
	sucmsg=sucmsg+"<br><li>题目清理完毕! 本次操作执行时间："&FormatNumber((timer()-startime1)*1000,3)&" 毫秒"
	call suc()	
end sub
%>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>