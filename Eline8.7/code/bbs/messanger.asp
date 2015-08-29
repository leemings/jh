<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="z_visual_const.asp"-->
<html>
<head>
<title><%=Forum_info(0)%>--短消息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>
<body <%=Forum_body(11)%>" onkeydown="if(event.keyCode==13 && event.ctrlKey && document.messager=='[object]')messager.submit()">
<%dim msg
dim abgcolor
dim username
if not founduser then
	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	founderr=true
end if
if Cint(GroupSetting(32))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有浏览发送短信的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	founderr=true
end if
stats="短信管理"
if founderr then
	call dvbbs_error()
else
	select case request("action")
	case "new"
		call sendmsg()
	case "read"
		call read()
	case "outread"
		call read()
	case "delet"
		call delete()
	case "send"
		call savemsg()
	case "fw"
		call fw()
	case "edit"
		call edit()
	case "savedit"
		call savedit()
	case "删除收件"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "的控制面板","usermanager.asp")
		call delinbox()
	case "清空收件箱"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "的控制面板","usermanager.asp")
		call AllDelinbox()
	case "删除草稿"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "的控制面板","usermanager.asp")
		call deloutbox()
	case "清空草稿箱"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "的控制面板","usermanager.asp")
		call AllDeloutbox()
	case "删除已发送的消息"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "的控制面板","usermanager.asp")
		call delissend()
	case "清空已发送的消息"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "的控制面板","usermanager.asp")
		call AllDelissend()
	case "删除垃圾"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "的控制面板","usermanager.asp")
		call delrecycle()
	case "清空垃圾箱"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "的控制面板","usermanager.asp")
		call AllDelrecycle()
	case else
  	errmsg=errmsg+"<br>"+"<li>请指定正确的参数。"
		founderr=true
	end select
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

'发送信息
sub sendmsg()
	stats="发送短信"
	dim sendtime,title,content
	if request("id")<>"" and isNumeric(request("id")) then
		set rs=server.createobject("adodb.recordset")
		sql="select sendtime,title,content from message where incept='"&membername&"' and id="&request("id")
		rs.open sql,conn,1,1
		if not(rs.eof and rs.bof) then
			sendtime=rs("sendtime")
			title="RE " & rs("title")
			content=rs("content")
		end if
		rs.close
		set rs=nothing
	end if%>
	<form action="messanger.asp" method=post name=messager>
	<input type=hidden name="action" value="send">
	<table cellpadding=3 cellspacing=1 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
		<tr> 
			<th colspan=2>发送短消息（请输入完整信息）</th>
		</tr>
		<tr> 
			<td class=tablebody1 valign=middle><b>收件人：</b></td>
			<td class=tablebody1 valign=middle><input type=text name="touser" value="<%=request("touser")%>" size=50> <SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)><OPTION selected value="">选择</OPTION><%
				set rs=server.createobject("adodb.recordset")
				sql="select F_friend from Friend where F_username='"&membername&"' and F_type=0 order by F_addtime desc"
				rs.open sql,conn,1,1
				do while not rs.eof
					%><OPTION value="<%=rs(0)%>"><%=rs(0)%></OPTION><%
					rs.movenext
				loop
				rs.close
				set rs=nothing
			%></SELECT></td>
		</tr>
		<tr> 
			<td class=tablebody1 valign=top width=15%><b>标题：</b></td>
			<td class=tablebody1 valign=middle><input type=text name="title" size=70 maxlength=80 value="<%=title%>"></td>
		</tr>
		<tr> 
			<td class=tablebody1 valign=top width=15%><b>内容：</b></td>
			<td class=tablebody1 valign=middle><textarea cols=70 rows=6 name="message" title="Ctrl+Enter发送"><%
				if request("id")<>"" then
					response.write "====== 在 "&sendtime&" 您来信中写道： ======"
					response.write vbNewLine
					response.write server.htmlencode(content)
					response.write vbNewLine
					response.write "====================================="
				end if
			%></textarea></td>
		</tr>
		<tr> 
			<td  class=tablebody1 colspan=2><b>说明</b>：<br>① 您可以使用<b>Ctrl+Enter</b>键快捷发送短信<br>② 可以用英文状态下的逗号将用户名隔开实现群发，最多<b><%=GroupSetting(33)%></b>个用户<br>③ 标题最多<b>50</b>个字符，内容最多<b><%=GroupSetting(34)%></b>个字符<br><font color=<%=forum_body(8)%>>④ 每发给一个收件人收取费用<%
				if isnull(myvip) or myvip<>1 then
					response.write forum_user(26)
				else
					response.write forum_user(27)
				end if
			%>元，您现在有现金<%=mymoney%>元</font></td>
		</tr>
		<tr> 
			<td  class=tablebody2 valign=middle colspan=2 align=center><input type=Submit value="发送" name=Submit>&nbsp; <input type=Submit value="保存" name=Submit>&nbsp; <input type="reset" name="Clear" value="清除">&nbsp; <%
				if request("reaction")="chatlog" then
					%><input type=button value="关闭聊天记录" name="chatlog" onclick="location.href='?action=new&id=<%=request("id")%>&touser=<%=request("touser")%>'"><%
				else
					%><input type=button value="查看聊天记录" name="chatlog" onclick="location.href='?action=new&id=<%=request("id")%>&touser=<%=request("touser")%>&reaction=chatlog'"><%
				end if
			%>&nbsp; <input type="button" name="close" value="关闭" onclick="window.close()"></td>
		</tr>
		<%if request("reaction")="chatlog" then%>
			<tr> 
				<th colspan=2>我与<%=request("touser")%>的聊天记录</th>
			</tr>
			<%if membername=request("touser") then%>
				<tr> 
					<td class=tablebody1 colspan=2>自己跟自己的聊天记录没什么好看的：）</td>
				</tr>
			<%else
				set rs=server.createobject("adodb.recordset")
				sql="select * from message where ((sender='"&trim(membername)&"' and incept='"&replace(request("touser"),"'","")&"') or (sender='"&replace(request("touser"),"'","")&"' and incept='"&membername&"')) and delS=0 order by sendtime desc"
				rs.open sql,conn,1,1
				if rs.eof and rs.bof then
					%><tr> 
						<td class=tablebody1 colspan=2>还没有任何聊天记录！</td>
					</tr><%
				else
					do while not rs.eof
						%><tr>
							<td class=tablebody2 height=25 colspan=2><%
								if rs("sender")=membername then
									%>在<b><%=rs("sendtime")%></b>，您发送此消息给<b><%=htmlencode(rs("incept"))%></b>！<%
								else
									%>在<b><%=rs("sendtime")%></b>，<b><%=htmlencode(rs("sender"))%></b>给您发送的消息！<%
								end if
							%></td>
						</tr>
						<tr>
							<td  class=tablebody1 valign=top align=left colspan=2><b>消息标题：<%=htmlencode(rs("title"))%></b><hr size=1><%=dvbcode(rs("content"),4,3)%></td>
						</tr>
						<%rs.movenext
					loop
				end if
				rs.close
				set rs=nothing
			end if
		end if%>
	</table>
	</form>
<%end sub
'读取信息
sub read()
	stats="阅读短信"
	if request("id")="" or not isNumeric(request("id")) then
		Errmsg=Errmsg+"<br>"+"<li>请指定相关参数。"
		Founderr=true
		exit sub
	end if
	set rs=server.createobject("adodb.recordset")
	if request("action")="read" then
		sql="update message set flag=1 where ID="&cstr(request("id"))
		conn.execute(sql)
	end if
	sql="select * from message where (incept='"&membername&"' or sender='"&membername&"') and id="&cstr(request("id"))
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		errmsg=errmsg+"<br>"+"<li>你是不是跑到别人的信箱啦、或者该信息已经收件人删除。"
		founderr=true
	end if
	if not founderr then%>
		<table cellpadding=3 cellspacing=1 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
			<tr>
				<th colspan=2 height=25>欢迎使用短消息接收，<%=membername%></th>
			</tr>
			<tr>
				<td class=tablebody1 valign=middle align=center width=140 rowspan=4><%
					if lcase(trim(rs("sender")))=lcase(trim(membername)) then
						call ShowUserVisual(rs("incept"),140,"tablebody1",false)
					else
						call ShowUserVisual(rs("sender"),140,"tablebody1",false)
					end if
				%></td>
				<td class=tablebody1 valign=middle align=center height=25><a href="messanger.asp?action=delet&id=<%=rs("id")%>"><img src="<%=Forum_info(7)%>m_delete.gif" border=0 alt="删除消息"></a> &nbsp; <a href="messanger.asp?action=new"><img src="<%=Forum_info(7)%>m_write.gif" border=0 alt="发送消息"></a> &nbsp;<a href="messanger.asp?action=new&touser=<%=htmlencode(rs("sender"))%>&id=<%=request("id")%>"><img src="<%=Forum_info(7)%>m_reply.gif" border=0 alt="回复消息"></a>&nbsp;<a href="messanger.asp?action=fw&id=<%=request("id")%>"><img src=<%=Forum_info(7)%>m_fw.gif border=0 alt=转发消息></a></td>
			</tr>
			<tr>
				<td class=tablebody2 height=25><%
					if request("action")="outread" then
						%>在<b><%=rs("sendtime")%></b>，您发送此消息给<b><%=htmlencode(rs("incept"))%></b>！<%
					else
						%>在<b><%=rs("sendtime")%></b>，<b><%=htmlencode(rs("sender"))%></b>给您发送的消息！<%
					end if
				%></td>
			</tr>
			<tr>
				<td  class=tablebody1 valign=top align=left><b>消息标题：<%=htmlencode(rs("title"))%></b><hr size=1><%=dvbcode(rs("content"),4,3)%></td>
			</tr>
			<%rs.close
			set rs=nothing
			sql="select id,sender from message where incept='"&membername&"' and flag=0 and issend=1 and id>"&cstr(request("id")&" order by sendtime")
			set rs=conn.execute(sql)
			if not (rs.eof and rs.bof) then%>
				<tr>
					<td  class=tablebody2 valign=top align=right><a href=messanger.asp?action=read&id=<%=rs(0)%>&sender=<%=rs(1)%>>[读取下一条信息]</a></td>
				</tr>
			<%end if
			rs.close
			set rs=nothing%>
		</table>
	<%end if
end sub
'转发信息
sub fw()
	stats="发送短信"
	dim title,content,sender
	if request("id")<>"" and isNumeric(request("id")) then
		set rs=server.createobject("adodb.recordset")
		sql="select title,content,sender from message where (incept='"&membername&"' or sender='"&membername&"') and id="&request("id")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			Errmsg=Errmsg+"<br>"+"<li>请选择相关参数。"
			Founderr=true
			exit sub
		else
			title=rs("title")
			content=rs("content")
			sender=rs("sender")
		end if
		rs.close
		set rs=nothing
	end if%>
	<form action="messanger.asp" method=post name=messager>
	<table cellpadding=3 cellspacing=1 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
		<tr> 
	  	<th colspan=2 height=25><input type=hidden name="action" value="send">发送短消息--请完整输入下列信息</th>
		</tr>
		<tr> 
	  	<td class=tablebody1 valign=middle width=15%><b>收件人：</b></td>
	  	<td  class=tablebody1 valign=middle><input type=text name="touser" value="<%=request("touser")%>" size=50> <SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)><OPTION selected value="">选择</OPTION><%
				set rs=server.createobject("adodb.recordset")
				sql="select F_friend from Friend where F_username='"&membername&"' and F_type=0 order by F_addtime desc"
				rs.open sql,conn,1,1
				do while not rs.eof
					%><OPTION value="<%=rs(0)%>"><%=rs(0)%></OPTION><%
					rs.movenext
				loop
				rs.close
				set rs=nothing
			%></SELECT></td>
		</tr>
		<tr> 
			<td class=tablebody1 valign=top><b>标题：</b></td>
			<td class=tablebody1 valign=middle><input type=text name="title" size=70 maxlength=80 value="Fw：<%=title%>">&nbsp;</td>
		</tr>
		<tr> 
			<td class=tablebody1 valign=top><b>内容：</b></td>
			<td class=tablebody1 valign=middle><textarea cols=70 rows=6 name="message" title="Ctrl+Enter发送"><%
				response.write "========== 下面是转发信息 ========="
				response.write vbNewLine
				response.write "原发件人："&sender&vbNewLine&vbNewLine
				response.write server.htmlencode(content)
				response.write "==================================="
			%></textarea></td>
		</tr>
		<tr> 
			<td  class=tablebody1 colspan=2><b>说明</b>：<br>① 您可以使用<b>Ctrl+Enter</b>键快捷发送短信<br>② 可以用英文状态下的逗号将用户名隔开实现群发，最多<b><%=GroupSetting(33)%></b>个用户<br>③ 标题最多<b>50</b>个字符，内容最多<b><%=GroupSetting(34)%></b>个字符<br><font color=<%=forum_body(8)%>>④ 每发给一个收件人收取费用<%
				if isnull(myvip) or myvip<>1 then
					response.write forum_user(26)
				else
					response.write forum_user(27)
				end if
			%>元，您现在有现金<%=mymoney%>元</font></td>
		</tr>
		<tr> 
			<td class=tablebody2 valign=middle colspan=2 align=center> <input type=Submit value="发送" name=Submit>&nbsp; <input type=Submit value="保存" name=Submit>&nbsp; <input type="reset" name="Clear" value="清除">&nbsp; <input type="button" name="close" value="关闭" onclick="window.close()"></td>
		</tr>
	</table>
	</form>
<%end sub

sub savemsg()
	stats="发送短信成功"
	dim incept,title,message,subtype
	if request("touser")="" then
		errmsg=errmsg+"<br>"+"<li>您忘记填写发送对象了吧。"
		founderr=true
		exit sub
	else
		incept=CheckStr(request("touser"))
		incept=split(incept,",")
	end if
	if request("title")="" then
		errmsg=errmsg+"<br>"+"<li>您还没有填写标题呀。"
		founderr=true
		exit sub
	elseif strlength(request("title"))>50 then
		errmsg=errmsg+"<br>"+"<li>标题限定最多50个字符。"
		founderr=true
		exit sub
	else
		title=CheckStr(request("title"))
	end if
	if request("message")="" then
		errmsg=errmsg+"<br>"+"<li>内容是必须要填写的噢。"
		founderr=true
		exit sub
	elseif strlength(request("message"))>Cint(GroupSetting(34)) then
		errmsg=errmsg+"<br>"+"<li>内容限定最多"&GroupSetting(34)&"个字符。"
		founderr=true
		exit sub
	else
		message=CheckStr(request("message"))
	end if
	if isnull(myvip) or myvip<>1 then
		if mymoney<(ubound(incept)+1)*clng(forum_user(26)) then
			errmsg=errmsg+"<br>"+"<li>您的现金不够给这些人发送短信。"
			founderr=true
			exit sub
		end if
	else
		if mymoney<(ubound(incept)+1)*clng(forum_user(27)) then
			errmsg=errmsg+"<br>"+"<li>您的现金不够给这些人发送短信。"
			founderr=true
			exit sub
		end if
	end if
	for i=0 to ubound(incept)
		sql="select username,msgflag from [user] where username='"&replace(incept(i),"'","")&"'"
		set rs=server.createobject("ADODB.Recordset")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			errmsg=errmsg+"<br>"+"<li>论坛没有用户“"&replace(incept(i),"'","")&"”，看看你的发送对象写对了嘛？"
			founderr=true
			exit sub
		elseif not isnull(rs("msgFlag")) and rs("msgflag")=9 then
			errmsg=errmsg+"<br>"+"<li>用户“"&replace(incept(i),"'","")&"”目前不接收任何短消息！"
			founderr=true
			exit sub
		elseif not isnull(rs("msgFlag")) and rs("msgflag")=1 then
			if conn.execute("select top 1 f_friend from friend where f_username='"&replace(incept(i),"'","")&"' and f_friend='"&membername&"'").eof then
				errmsg=errmsg+"<br>"+"<li>用户“"&replace(incept(i),"'","")&"”目前只接收好友的短消息！"
				founderr=true
				exit sub
			end if
		end if
		dim rst
		set rst=conn.execute("select top 1 F_id from [Friend] where F_username='"&replace(incept(i),"'","")&"' and F_friend='"&membername&"' and F_type=1")
		if not (rst.eof and rst.bof) then
			errmsg=errmsg+"<br>"+"<li>您已经被"&rs(0)&"屏蔽了，您不能给他发送短信！"
			founderr=true
			exit sub
		end if
		rst.close
		set rs=nothing
		if request("Submit")="发送" then
			sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,1)"
			subtype="已发送信息"
			if isnull(myvip) or myvip<>1 then
				conn.execute("update [user] set userwealth=userwealth-"&forum_user(26)&" where username='"&membername&"'")
			else
				conn.execute("update [user] set userwealth=userwealth-"&forum_user(27)&" where username='"&membername&"'")
			end if
		elseif request("Submit")="保存" then
			sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,0)"
			subtype="发件箱"
		else
			sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,1)"
			subtype="已发送信息"
			if isnull(myvip) or myvip<>1 then
				conn.execute("update [user] set userwealth=userwealth-"&forum_user(26)&" where username='"&membername&"'")
			else
				conn.execute("update [user] set userwealth=userwealth-"&forum_user(27)&" where username='"&membername&"'")
			end if
		end if
		conn.execute(sql)
		if i>Cint(GroupSetting(33))-1 then
			errmsg=errmsg+"<br>"+"<li>最多只能发送给"&GroupSetting(33)&"个用户，您的名单"&GroupSetting(33)&"位以后的请重新发送"
			founderr=true
			exit sub
			exit for
		end if
	next
	sucmsg=sucmsg+"<br>"+"<li><b>恭喜您，发送短信息成功。</b><br>发送的消息同时保存在您的"&subtype&"中。"
	call dvbbs_suc()
end sub

'更改信息
sub edit()
	stats="修改短信"
	dim incept,title,content,id
	if request("id")<>"" and isNumeric(request("id")) then
		set rs=server.createobject("adodb.recordset")
		sql="select id,incept,title,content from message where sender='"&membername&"' and issend=0 and id="&request("id")
		rs.open sql,conn,1,1
		if not(rs.eof and rs.bof) then
			incept=rs("incept")
			title=rs("title")
			content=rs("content")
			id=rs("id")
		else
			Errmsg=Errmsg+"<br>"+"<li>没有找到您要编辑的信息。"
			Founderr=true
			exit sub
		end if
		rs.close
		set rs=nothing
	else
		Errmsg=Errmsg+"<br>"+"<li>请指定相关参数。"
		Founderr=true
		exit sub
	end if%>
	<form action="messanger.asp" method=post name=messager>
	<table cellpadding=3 cellspacing=1 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
		<tr> 
			<th colspan=2 height=25><input type=hidden name="action" value="savedit"><input type=hidden name="id" value="<%=id%>">发送短消息--请完整输入下列信息</th>
		</tr>
		<tr> 
			<td  class=tablebody1 valign=middle width=70><b>收件人：</b></td>
			<td  class=tablebody1 valign=middle><input type=text name="touser" value="<%=incept%>" size=70></td>
		</tr>
		<tr> 
			<td  class=tablebody1 valign=top><b>标题：</b></td>
			<td  class=tablebody1 valign=middle><input type=text name="title" size=70 maxlength=80 value="<%=title%>"></td>
		</tr>
		<tr> 
			<td  class=tablebody1 valign=top><b>内容：</b></td>
			<td  class=tablebody1 valign=middle><textarea cols=70 rows=8 name="message" title=""><%=server.htmlencode(content)%></textarea></td>
		</tr>
		<tr> 
			<td  class=tablebody1 colspan=2><b>说明</b>：<br>① 您可以使用<b>Ctrl+Enter</b>键快捷发送短信<br>② 标题最多<b>50</b>个字符，内容最多<b><%=GroupSetting(34)%></b>个字符<BR><font color=<%=forum_body(8)%>>③ 每发给一个收件人收取费用<%
				if isnull(myvip) or myvip<>1 then
					response.write forum_user(26)
				else
					response.write forum_user(27)
				end if
			%>元，您现在有现金<%=mymoney%>元</font></td>
		</tr>
		<tr> 
			<td  class=tablebody2 valign=middle colspan=2 align=center><input type=Submit value="发送" name=Submit>&nbsp; <input type=Submit value="保存" name=Submit>&nbsp; <input type="reset" name="Clear" value="清除">&nbsp; <input type="button" name="close" value="关闭" onclick="window.close()"></td>
		</tr>
	</table>
	</form>
<%end sub

sub savedit()
	dim incept,title,message,subtype
	if request("id")="" or not isNumeric(request("id")) then
		Errmsg=Errmsg+"<br>"+"<li>请指定相关参数。"
		Founderr=true
		exit sub
	end if
	if request("touser")="" then
		errmsg=errmsg+"<br>"+"<li>您忘记填写发送对象了吧。"
		founderr=true
		exit sub
	else
		incept=checkStr(request("touser"))
	end if
	if request("title")="" then
		errmsg=errmsg+"<br>"+"<li>您还没有填写标题呀。"
		founderr=true
		exit sub
	else
		title=checkStr(request("title"))
	end if
	if request("message")="" then
		errmsg=errmsg+"<br>"+"<li>内容是必须要填写的噢。"
		founderr=true
		exit sub
	else
		message=checkStr(request("message"))
	end if
	sql="select username from [user] where username='"&incept&"'"
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		errmsg=errmsg+"<br>"+"<li>论坛没有用户"&incept&"，看看你的发送对象写对了嘛？"
		founderr=true
		exit sub
	end if
	dim rst
	set rst=conn.execute("select top 1 F_id from [Friend] where F_username='"&incept&"' and F_friend='"&membername&"' and F_type=1")
	if not (rst.eof and rst.bof) then
		errmsg=errmsg+"<br>"+"<li>您已经被"&rs(0)&"屏蔽了，您不能给他发送短信！"
		founderr=true
		exit sub
	end if
	rst.close
	set rs=nothing
	if request("Submit")="发送" then
		if isnull(myvip) or myvip<>1 then
			if mymoney<clng(forum_user(26)) then
				errmsg=errmsg+"<br>"+"<li>您的现金不够发送这封短信。"
				founderr=true
				exit sub
			end if
		else
			if mymoney<clng(forum_user(27)) then
				errmsg=errmsg+"<br>"+"<li>您的现金不够发送这封短信。"
				founderr=true
				exit sub
			end if
		end if
		sql="update message set incept='"&incept&"',sender='"&membername&"',title='"&title&"',content='"&message&"',sendtime=Now(),flag=0,issend=1 where id="&request("id")
		subtype="已发送信息"
		if isnull(myvip) or myvip<>1 then
			conn.execute("update [user] set userwealth=userwealth-"&forum_user(26)&" where username='"&membername&"'")
		else
			conn.execute("update [user] set userwealth=userwealth-"&forum_user(27)&" where username='"&membername&"'")
		end if
	else
		sql="update message set incept='"&incept&"',sender='"&membername&"',title='"&title&"',content='"&message&"',sendtime=Now(),flag=0,issend=0 where id="&request("id")
		subtype="发件箱"
	end if
	conn.execute(sql)
	sucmsg=sucmsg+"<br>"+"<li><b>恭喜您，发送短信息成功。</b><br>发送的消息同时保存在您的"&subtype&"中。"
	call dvbbs_suc()
end sub

'收件逻辑删除，置于回收站，入口字段delR，可用于批量及单个删除
sub delinbox()
	dim delid
	delid=replace(request("id"),"'","")
	delid=replace(delid,";","")
	delid=replace(delid,"--","")
	delid=replace(delid,")","")
	if delid="" or isnull(delid) then
	Errmsg=Errmsg+"<li>"+"请选择相关参数。"
	Founderr=true
	else
		conn.execute("update message set delR=1 where incept='"&trim(membername)&"' and id in ("&delid&")")
		sucmsg=sucmsg+"<br>"+"<li>删除短信息成功。删除的消息将转移到您的回收站。"
		call dvbbs_suc()
	end if
end sub

sub AllDelinbox()
	conn.execute("update message set delR=1 where incept='"&trim(membername)&"' and delR=0")
	sucmsg=sucmsg+"<br>"+"<li>删除短信息成功。删除的消息将转移到您的回收站。"
	call dvbbs_suc()
end sub

'发件逻辑删除，置于回收站，入口字段delS，可用于批量及单个删除
sub deloutbox()
	dim delid
	delid=replace(request("id"),"'","")
	delid=replace(delid,";","")
	delid=replace(delid,"--","")
	delid=replace(delid,")","")
	if delid="" or isnull(delid) then
		Errmsg=Errmsg+"<li>"+"请选择相关参数。"
		Founderr=true
	else
		conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and issend=0 and id in ("&delid&")")
		sucmsg=sucmsg+"<br>"+"<li>删除短信息成功。删除的消息将转移到您的回收站。"
		call dvbbs_suc()
	end if
end sub

sub AllDeloutbox()
	conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and delS=0 and issend=0")
	sucmsg=sucmsg+"<br>"+"<li>删除短信息成功。删除的消息将转移到您的回收站。"
	call dvbbs_suc()
end sub

'已发送逻辑删除，置于回收站，入口字段delS，可用于批量及单个删除
'delS：0未操作，1发送者删除，2发送者从回收站删除
sub delissend()
	dim delid
	delid=replace(request("id"),"'","")
	delid=replace(delid,";","")
	delid=replace(delid,"--","")
	delid=replace(delid,")","")
	if delid="" or isnull(delid) then
		Errmsg=Errmsg+"<li>"+"请选择相关参数。"
		Founderr=true
	else
		conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and issend=1 and id in ("&delid&")")
		sucmsg=sucmsg+"<br>"+"<li>删除短信息成功。删除的消息将转移到您的回收站。"
		call dvbbs_suc()
	end if
end sub

sub AllDelissend()
	conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and delS=0 and issend=1")
	sucmsg=sucmsg+"<br>"+"<li>删除短信息成功。删除的消息将转移到您的回收站。"
	call dvbbs_suc()
end sub

'用户能完全删除收到信息和逻辑删除所发送信息，逻辑删除所发送信息设置入口字段delS参数为2
sub delrecycle()
	dim delid
	delid=replace(request("id"),"'","")
	delid=replace(delid,";","")
	delid=replace(delid,"--","")
	delid=replace(delid,")","")
	if delid="" or isnull(delid) then
		Errmsg=Errmsg+"<li>"+"请选择相关参数。"
		Founderr=true
		exit sub
	else
		conn.execute("delete from message where incept='"&membername&"' and delR=1 and id in ("&delid&")")
		conn.execute("update message set delS=2 where sender='"&trim(membername)&"' and delS=1 and id in ("&delid&")")
		sucmsg=sucmsg+"<br>"+"<li>删除短信息成功。删除的消息将不可恢复。"
		call dvbbs_suc()
	end if
end sub

sub AllDelrecycle()
	conn.execute("delete from message where incept='"&membername&"' and delR=1")	
	conn.execute("update message set delS=2 where sender='"&trim(membername)&"' and delS=1")
	sucmsg=sucmsg+"<br>"+"<li>删除短信息成功。删除的消息将不可恢复。"
	call dvbbs_suc()
end sub

sub delete()
	dim delid
	delid=checkstr(request("id"))
	if not isNumeric(request("id")) or delid="" or isnull(delid) then
		Errmsg=Errmsg+"<li>"+"请选择相关参数。"
		Founderr=true
	else
		conn.execute("update message set delR=1 where incept='"&trim(membername)&"' and id="&delid)
		conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and id="&delid)
		sucmsg=sucmsg+"<br>"+"<li>删除短信息成功。删除的消息将置于您的回收站内。"
		call dvbbs_suc()
	end if
end sub%>
<script language="javascript"> 
function DoTitle(addTitle) {  
 var revisedTitle;  
 var currentTitle = document.messager.touser.value; 

 if(currentTitle=="") revisedTitle = addTitle; 
 else { 
  var arr = currentTitle.split(","); 
  for (var i=0; i < arr.length; i++) { 
   if( addTitle.indexOf(arr[i]) >=0 && arr[i].length==addTitle.length ) return; 
  } 
  revisedTitle = currentTitle+","+addTitle; 
 } 

 document.messager.touser.value=revisedTitle;  
 document.messager.touser.focus(); 
 return; 
} 
</script>