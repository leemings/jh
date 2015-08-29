<!--#include file="conn.asp"-->
<!--#include file="z_testconn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'--------------------
'edit by 绿水青山
'--------------------

	dim ars
	
	stats="开心辞典 审核上传题目"
	
	if not master then
		errmsg=errmsg+"<br><li>您没有在开心辞典审核上传题的权限，请确认你已经登录并且具有相应的权限！"
		founderr=true
		stats="开心辞典 错误信息"
		call nav()
		call head_var(0,0,"开心辞典","z_test.asp")		
		call dvbbs_error()
	else
		call nav()
		call head_var(0,0,"开心辞典","z_test.asp")
		
		select case request("action")
			case "Auditing"
				call Auditing()
			case "通过审核"
				call pass()
			case "删除本题"
				call del()
			case "passall"
				call passall()
			case "delall"
				call delall()
			case else		
				call UploadList()
		end select
	end if
	set aconn=nothing
	call activeonline()
	call footer()
'================================================
'-----------------上传题目列表--------------
sub UploadList()
%>
<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
	<tr><th>欢迎 <%=membername%> 进入开心辞典 管理中心</th></tr>
	<tr><td class=tablebody2>
	<%call Admin_Head()%>
		<table align=center cellspacing=1 cellpadding=3 class=tableborder1 style="width:97%">
		  <tr><th colspan="7">上传题目管理</th></tr> 	
		  <tr>
			<td width="60" align="center" class=TopLighNav1><b>题　号</b></td>
			<td width="*" align="center" class=TopLighNav1><b>题　目</b></td>
			<td width="100" align="center" class=TopLighNav1><b>题目类别</b></td>
			<td width="50" align="center" class=TopLighNav1 ><b>题型</b></td>
			<td width="90" align="center" class=TopLighNav1><b>本题上传者</b></td>
			<td width="65" align="center" class=TopLighNav1><b>上传时间</b></td> 
			<td width="130" align="center" class=TopLighNav1><b>操作</b></td> 
		  </tr>
<%
			dim HaveUpLoad
			HaveUpLoad=true

			set ars=aconn.execute("select * from [testtemp] order by AddTime desc")
			if ars.eof and ars.bof then
				HaveUpLoad=false
				response.Write "<tr><td colspan=7 class=tablebody1>暂时没有上传题目需要审核</td></tr>"
			else
				do while not ars.eof
%>
				  <tr>
					<td align="center" class=tablebody2><%=ars("id")%></td>
					<td class=tablebody1><a href="?action=Auditing&lbid=<%=ars("lb")%>&id=<%=ars("id")%>"><%=ars("title")%></a></td>
					<td class=tablebody1 align="center"><%=GetLbName(ars("lb"))%></td>
					<td class=tablebody1 align="center"><%if ars("tx")=1 then%>填空题<%else%>选择题<%end if%></td> 
					<td class=tablebody1 align="center"><%=ars("username")%></td>
					<td class=tablebody1 align="center"><%=ars("addtime")%></td> 
					<td class=tablebody1 align="center"><a href="?action=Auditing&lbid=<%=ars("lb")%>&id=<%=ars("id")%>" title="进入审核">审核</a> | <a href="?action=通过审核&flag=auto&id=<%=ars("id")%>" title="直接通过审核" onclick="javascript:{if(confirm('您确定不查看具体题目内容就通过本题的审核吗?')){return true;}return false;}">通过</a> | <a href="?action=删除本题&id=<%=ars("id")%>" title="直接删除本题" onclick="javascript:{if(confirm('您确定删除本题吗?')){return true;}return false;}">删除</a></td>
				  </tr>
<%				
					ars.movenext
				loop	
			end if
			ars.close
			set ars=nothing	
%>		  
		</table> 
		<br>
	</td></tr>	
	<tr><td class=tablebody1 align="center"><%if HaveUpLoad then%><a href="?action=passall" title="直接通过所有待审核的题目" onclick="javascript:{if(confirm('您确定不查看具体题目内容就通过所有待审核的题目吗?')){return true;}return false;}">全部通过</a> | <a href="?action=delall" title="直接删除所有待审核的题目" onclick="javascript:{if(confirm('您确定删除所有待审核的题目吗?')){return true;}return false;}">全部删除</a> | <%end if%><a href="Z_test.asp">返  回</a></td></tr>
</table>	
<%	
end sub
'------------审核题目--------------
sub Auditing()
	dim id,lbid
	id=request("id")
	if not isInteger(id) then
		errmsg=errmsg+"<br><li>错误的题目参数"
		call test_err()
		exit sub
	end if

	sql="select * from [testtemp] where id="&id
	set ars=aconn.execute(sql)
	if ars.eof and not ars.bof then
		errmsg=errmsg+"<br><li>没有找到相应的题目"
		founderr=true
		ars.close
		call test_err()
		exit sub
	end if
	lbid=request("lbid")
%>
	<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
	<form method="POST" action="Z_TestAdminUpLoad.asp">
	<tr><th>欢迎 <%=membername%> 进入开心辞典 管理中心</th></tr>
	<tr><td class=tablebody2>
<%call Admin_Head()%>
<table align=center cellspacing=1 cellpadding=3 class=tableborder1 style="width:97%">
  <tr><th colspan=8>审核上传题目</th></tr> 
  <tr>
	<td width="70" align="center" class=tablebody2><b>临时题号</b></td>
	<td width="82" class=tablebody1>&nbsp;<%=ars("id")%></td>
	<td width="98" align="center" class=tablebody2><b>题目所属类别</b></td>
	<td width="102" class=tablebody1><select name="lb"><%
	set rs=aconn.execute("select lb,id from [testlb]")
	do while not rs.eof
%>
		<option value=<%=rs(1)%> <%if cint(rs(1))=cint(lbid) then%> selected <%end if%>><%=rs(0)%></option>
<%
  		rs.movenext
	loop
	rs.close
	set rs=nothing
%></select></td>
	<td width="57" align="center" class=tablebody2 ><b>题型</b></td>
	<td width="63" class=tablebody1><%if ars("tx")=1 then%>填空题<%else%>选择题<%end if%></td>
	<td width="90" align="center" class=tablebody2><b>本题上传者</b></td>
    <td width="102" class=tablebody1><%=ars("username")%></td>
  </tr>
  <tr>
    <td width="70" class=tablebody2 align="center"><b>题&nbsp;&nbsp;目</b></td>
    <td width="600" colspan="7" align="left" class=tablebody1>&nbsp;<textarea rows="12" name="title" cols="70"><%=ars("title")%></textarea>
	</td>
  </tr>
  <tr>
    <td width="70"  align="center" class=tablebody2><%if ars("tx")=1 then%>正确答案：<%else%>答案 1：<%end if%></td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key1" size="70" value='<%=ars("key1")%>'></td>
  </tr>
 <%if ars("tx")=0 then%>
  <tr>
    <td width="70" align="center" class=tablebody2>答案 2：</td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key2" size="70"  value='<%=ars("key2")%>'></td>
  </tr>
  <tr>
    <td width="70" align="center" class=tablebody2>答案 3：</td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key3" size="70" value='<%=ars("key3")%>'></td>
  </tr>
  <tr>
    <td width="70"  align="center" class=tablebody2>答案 4：</td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key4" size="70"  value='<%=ars("key4")%>'></td>
  </tr>
 <%end if%>
  <tr>
    <td width="70" class=tablebody2 align="center"><%if ars("tx")=0 then%>正确答案代号：<%end if%></td>
    <td width="81" class=tablebody1>&nbsp;<%if ars("tx")=0 then%><input type="text" name="ok" size="10"  value=<%=ars("ok")%>><%end if%></td>
    <td width="116" colspan="2" class=tablebody2> <b><span lang="zh-cn">限时</span></b>
      <input type="text" name="atime" size="7" value=<%=ars("timelimit")%>> <b><span lang="zh-cn"> 秒</span></td>
    <td width="172" colspan="2" class=tablebody2>
    <p align="center"><b>本题奖金 <input type="text" name="jj" size="7" value=<%=ars("jiangjin")%>> 元</b></td>
    <td width="231" colspan="2" class=tablebody1>
		<input type=hidden name="id" value=<%=ars("id")%>>
		<input type=hidden name="tx" value=<%=ars("tx")%>>
		<input type=hidden name="username" value=<%=ars("username")%>>
	</td>
  </tr>
</table><br>
<tr><td align="center" class=tablebody2>
	<input type=submit value="通过审核" name=action>　　<input type=submit value="删除本题" name=action>　　<input type="submit" value="返回题目列表" name="go">
</td></tr>
</form></table>   
<%
		set ars=nothing	
end sub

sub pass()
'write by 我来了 2003.1.1
'通过审核
	dim ausername
	dim id,lb,key(4),jj,atime,title,ok,tx
	id=request("id")
	if not isInteger(id) then
		errmsg=errmsg+"<br><li>错误的题目参数"
		call test_err()
		exit sub
	end if
	if request("flag")="auto" then		'直接审核题目
		set ars=aconn.execute("select * from [testtemp] where id="&id)
		if ars.eof and ars.bof then
			errmsg=errmsg+"<br><li>错误的题目参数，没有找到该题目"
			call test_err()
			ars.close
			exit sub
		end if
		lb=ars("lb"):		jj=ars("JiangJin"):		atime=ars("TimeLimit"):		title=ars("title")
		key(0)=ars("key1"):		key(1) =ars("key2"):		key(2) =ars("key3"):		key(3) =ars("key4")
		ok=ars("ok"):		tx=ars("tx"):		ausername=ars("username")
		ars.close
	else		'具体审核题目
		lb=request("lb"):		jj=request("jj"):		atime=request("atime"):		title=request("title")
		key(0)=request("key1"):		key(1) =request("key2"):		key(2) =request("key3"):		key(3) =request("key4")
		ok=request("ok"):		tx=request("tx"):		ausername=trim(replace(request("username"),"'",""))	
	end if	
	if lb="" or (not isnumeric(lb)) then
		errmsg=errmsg+"<br><li>错误的题目参数，请从有效连接进入"
		founderr=true
	end if
	if request("flag")<>"auto" then		'具体审核题目
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
		if cint(tx)=0 then		'是选择题就判断答案代码，填空题不用
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
	end if
	if ausername="" then ausername="管理员整理"
	if founderr then 
		call test_err()
		exit sub
	end if	
	set ars=server.createobject("adodb.recordset")	
	sql="select * from [test] where lb="&lb
	ars.open sql,aconn,1,3
	
	ars.addnew 
	ars("lb")=lb
	ars("jj")=jj
	ars("time")=atime
	ars("title")=title
	ars("key1")=key(0)
	ars("key2")=key(1) 
	ars("key3")=key(2) 
	ars("key4")=key(3) 
	ars("ok")=ok
	ars("tx")=tx
	ars("username")=ausername
	ars.update 
	ars.close 
	
	'删除临时表中的题目
	aconn.execute("delete from [testtemp] where id="&id)
	
	if ausername<>"管理员整理" then
		'取出上传题目的奖金和游戏币 
		dim jiangjin,auserupsinew
		set ars=aconn.execute("select userjl,userupsinew from [config]")
		jiangjin=ars(0)
		auserupsinew=ars(1)
		'把奖金和游戏币付给上传者 
		conn.execute ("update [user] set userWealth=userWealth+"&jiangjin&" where username='"&ausername&"'")   
		'更新上传用户的上传题目通过数和游戏币
		aconn.execute ("update testuser set userpass=userpass+1,usersinew=usersinew+"&auserupsinew&" where username='"&ausername&"'")
	
		'给上传者发送短信息通知
		dim message
		message="你上传的知识问答题目《"&request("title")&" 》已通过审核，同时，奖励给你"&jiangjin&"元社区元和"&auserupsinew&"个游戏币，请查收！"
		sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&ausername&"','"&membername&"','加入知识问答题目成功！','"&message&"',Now(),0,1)"
		conn.execute sql
	end if
	
	set ars=nothing	
	
	sucmsg=sucmsg+"<br><li>本题审核完成，请返回进行其他操作"
	rUrl="Z_TestAdminUpLoad.asp"
	call suc()
end sub

sub passall()
'write by 我来了  2003.1.1
'全部通过审核
	dim jiangjin,auserupsinew,message

	'取出上传题目的奖金和游戏币 
	set ars=aconn.execute("select userjl,userupsinew from [config]")
	jiangjin=ars(0)
	auserupsinew=ars(1)
	
	set ars=aconn.execute("select * from [testtemp]")
	if not (ars.eof and ars.bof) then 
		do while not ars.eof
			'把题目插入测试题库中 
			aconn.execute("insert into [test] (lb,jj,[time],title,key1,key2,key3,key4,ok,tx,username) values('"&ars("lb")&"',"&ars("JiangJin")&","&ars("TimeLimit")&",'"&ars("title")&"','"&ars("key1")&"','"&ars("key2")&"','"&ars("key3")&"','"&ars("key4")&"','"&ars("ok")&"',"&ars("tx")&",'"&ars("username")&"')") 
			'把奖金和游戏币付给上传者  
			conn.execute ("update [user] set userWealth=userWealth+"&jiangjin&" where username='"&ars("username")&"'")   
			'更新上传用户的上传题目通过数和游戏币
			aconn.execute ("update testuser set userpass=userpass+1,usersinew=usersinew+"&auserupsinew&" where username='"&ars("username")&"'")
			'给上传者发送短信息通知
			message="你录入的知识问答题目《"&ars("title")&"》没有通过审核，已从临时库中删除，请注意！"
			sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&ars("username")&"','"&membername&"','知识问答题目删除通知！','"&message&"',Now(),0,1)"
			conn.execute sql
			ars.movenext
		loop 
		'删除临时表中的题目
		aconn.execute("delete from [testtemp]") 
	end if
	ars.close
	set ars=nothing
	sucmsg=sucmsg+"<br><li>所有题目审核完成，请返回进行其他操作"
	rUrl="Z_TestAdminUpLoad.asp"
	call suc()	
end sub

sub del()
'write by 我来了  2003.1.1
'没有通过审核,题目从临时库中删除
	dim ausername,message
	ausername=replace(request("username"),"'","")
	'删除临时表中的题目
	aconn.execute("delete from [testtemp] where id="&request("id"))
	
	message="你录入的知识问答题目《"&request("title")&" 》没有通过审核，已从临时库中删除"
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&ausername&"','"&membername&"','知识问答题目删除通知！','"&message&"',Now(),0,1)"
	conn.execute sql
	
	sucmsg=sucmsg+"<br><li>题目已从临时库中删除，请返回进行其他操作"
	rUrl="Z_TestAdminUpLoad.asp"
	call suc()	
end sub

sub delall()
'write by 我来了 2003.1.1
'从临时库中删除所有待审核的上传的题目
	dim message 
	
	set ars=aconn.execute("select title,username from [testtemp]")
	if not (ars.eof and ars.bof) then 
		do while not ars.eof
			message="你录入的知识问答题目《"&ars(0)&"》没有通过审核，已从临时库中删除"
			sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&ars(1)&"','"&membername&"','知识问答题目删除通知！','"&message&"',Now(),0,1)"
			conn.execute sql
			ars.movenext
		loop 
		'删除临时表中的题目
		aconn.execute("delete from [testtemp]")
	end if
	ars.close	
	set ars=nothing	
	sucmsg=sucmsg+"<br><li>所有待审核的上传的题目已从临时库中删除，请返回进行其他操作"
	rUrl="Z_TestAdminUpLoad.asp"
	call suc()
end sub
%>