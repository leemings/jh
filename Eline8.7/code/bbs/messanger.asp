<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="z_visual_const.asp"-->
<html>
<head>
<title><%=Forum_info(0)%>--����Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>
<body <%=Forum_body(11)%>" onkeydown="if(event.keyCode==13 && event.ctrlKey && document.messager=='[object]')messager.submit()">
<%dim msg
dim abgcolor
dim username
if not founduser then
	errmsg=errmsg+"<br>"+"<li>��û��<a href=login.asp target=_blank>��¼</a>��"
	founderr=true
end if
if Cint(GroupSetting(32))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û��������Ͷ��ŵ�Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
	founderr=true
end if
stats="���Ź���"
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
	case "ɾ���ռ�"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "�Ŀ������","usermanager.asp")
		call delinbox()
	case "����ռ���"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "�Ŀ������","usermanager.asp")
		call AllDelinbox()
	case "ɾ���ݸ�"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "�Ŀ������","usermanager.asp")
		call deloutbox()
	case "��ղݸ���"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "�Ŀ������","usermanager.asp")
		call AllDeloutbox()
	case "ɾ���ѷ��͵���Ϣ"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "�Ŀ������","usermanager.asp")
		call delissend()
	case "����ѷ��͵���Ϣ"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "�Ŀ������","usermanager.asp")
		call AllDelissend()
	case "ɾ������"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "�Ŀ������","usermanager.asp")
		call delrecycle()
	case "���������"
		stats=request("action")
		call nav()
		call head_var(0,0,membername & "�Ŀ������","usermanager.asp")
		call AllDelrecycle()
	case else
  	errmsg=errmsg+"<br>"+"<li>��ָ����ȷ�Ĳ�����"
		founderr=true
	end select
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

'������Ϣ
sub sendmsg()
	stats="���Ͷ���"
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
			<th colspan=2>���Ͷ���Ϣ��������������Ϣ��</th>
		</tr>
		<tr> 
			<td class=tablebody1 valign=middle><b>�ռ��ˣ�</b></td>
			<td class=tablebody1 valign=middle><input type=text name="touser" value="<%=request("touser")%>" size=50> <SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)><OPTION selected value="">ѡ��</OPTION><%
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
			<td class=tablebody1 valign=top width=15%><b>���⣺</b></td>
			<td class=tablebody1 valign=middle><input type=text name="title" size=70 maxlength=80 value="<%=title%>"></td>
		</tr>
		<tr> 
			<td class=tablebody1 valign=top width=15%><b>���ݣ�</b></td>
			<td class=tablebody1 valign=middle><textarea cols=70 rows=6 name="message" title="Ctrl+Enter����"><%
				if request("id")<>"" then
					response.write "====== �� "&sendtime&" ��������д���� ======"
					response.write vbNewLine
					response.write server.htmlencode(content)
					response.write vbNewLine
					response.write "====================================="
				end if
			%></textarea></td>
		</tr>
		<tr> 
			<td  class=tablebody1 colspan=2><b>˵��</b>��<br>�� ������ʹ��<b>Ctrl+Enter</b>����ݷ��Ͷ���<br>�� ������Ӣ��״̬�µĶ��Ž��û�������ʵ��Ⱥ�������<b><%=GroupSetting(33)%></b>���û�<br>�� �������<b>50</b>���ַ����������<b><%=GroupSetting(34)%></b>���ַ�<br><font color=<%=forum_body(8)%>>�� ÿ����һ���ռ�����ȡ����<%
				if isnull(myvip) or myvip<>1 then
					response.write forum_user(26)
				else
					response.write forum_user(27)
				end if
			%>Ԫ�����������ֽ�<%=mymoney%>Ԫ</font></td>
		</tr>
		<tr> 
			<td  class=tablebody2 valign=middle colspan=2 align=center><input type=Submit value="����" name=Submit>&nbsp; <input type=Submit value="����" name=Submit>&nbsp; <input type="reset" name="Clear" value="���">&nbsp; <%
				if request("reaction")="chatlog" then
					%><input type=button value="�ر������¼" name="chatlog" onclick="location.href='?action=new&id=<%=request("id")%>&touser=<%=request("touser")%>'"><%
				else
					%><input type=button value="�鿴�����¼" name="chatlog" onclick="location.href='?action=new&id=<%=request("id")%>&touser=<%=request("touser")%>&reaction=chatlog'"><%
				end if
			%>&nbsp; <input type="button" name="close" value="�ر�" onclick="window.close()"></td>
		</tr>
		<%if request("reaction")="chatlog" then%>
			<tr> 
				<th colspan=2>����<%=request("touser")%>�������¼</th>
			</tr>
			<%if membername=request("touser") then%>
				<tr> 
					<td class=tablebody1 colspan=2>�Լ����Լ��������¼ûʲô�ÿ��ģ���</td>
				</tr>
			<%else
				set rs=server.createobject("adodb.recordset")
				sql="select * from message where ((sender='"&trim(membername)&"' and incept='"&replace(request("touser"),"'","")&"') or (sender='"&replace(request("touser"),"'","")&"' and incept='"&membername&"')) and delS=0 order by sendtime desc"
				rs.open sql,conn,1,1
				if rs.eof and rs.bof then
					%><tr> 
						<td class=tablebody1 colspan=2>��û���κ������¼��</td>
					</tr><%
				else
					do while not rs.eof
						%><tr>
							<td class=tablebody2 height=25 colspan=2><%
								if rs("sender")=membername then
									%>��<b><%=rs("sendtime")%></b>�������ʹ���Ϣ��<b><%=htmlencode(rs("incept"))%></b>��<%
								else
									%>��<b><%=rs("sendtime")%></b>��<b><%=htmlencode(rs("sender"))%></b>�������͵���Ϣ��<%
								end if
							%></td>
						</tr>
						<tr>
							<td  class=tablebody1 valign=top align=left colspan=2><b>��Ϣ���⣺<%=htmlencode(rs("title"))%></b><hr size=1><%=dvbcode(rs("content"),4,3)%></td>
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
'��ȡ��Ϣ
sub read()
	stats="�Ķ�����"
	if request("id")="" or not isNumeric(request("id")) then
		Errmsg=Errmsg+"<br>"+"<li>��ָ����ز�����"
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
		errmsg=errmsg+"<br>"+"<li>���ǲ����ܵ����˵������������߸���Ϣ�Ѿ��ռ���ɾ����"
		founderr=true
	end if
	if not founderr then%>
		<table cellpadding=3 cellspacing=1 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
			<tr>
				<th colspan=2 height=25>��ӭʹ�ö���Ϣ���գ�<%=membername%></th>
			</tr>
			<tr>
				<td class=tablebody1 valign=middle align=center width=140 rowspan=4><%
					if lcase(trim(rs("sender")))=lcase(trim(membername)) then
						call ShowUserVisual(rs("incept"),140,"tablebody1",false)
					else
						call ShowUserVisual(rs("sender"),140,"tablebody1",false)
					end if
				%></td>
				<td class=tablebody1 valign=middle align=center height=25><a href="messanger.asp?action=delet&id=<%=rs("id")%>"><img src="<%=Forum_info(7)%>m_delete.gif" border=0 alt="ɾ����Ϣ"></a> &nbsp; <a href="messanger.asp?action=new"><img src="<%=Forum_info(7)%>m_write.gif" border=0 alt="������Ϣ"></a> &nbsp;<a href="messanger.asp?action=new&touser=<%=htmlencode(rs("sender"))%>&id=<%=request("id")%>"><img src="<%=Forum_info(7)%>m_reply.gif" border=0 alt="�ظ���Ϣ"></a>&nbsp;<a href="messanger.asp?action=fw&id=<%=request("id")%>"><img src=<%=Forum_info(7)%>m_fw.gif border=0 alt=ת����Ϣ></a></td>
			</tr>
			<tr>
				<td class=tablebody2 height=25><%
					if request("action")="outread" then
						%>��<b><%=rs("sendtime")%></b>�������ʹ���Ϣ��<b><%=htmlencode(rs("incept"))%></b>��<%
					else
						%>��<b><%=rs("sendtime")%></b>��<b><%=htmlencode(rs("sender"))%></b>�������͵���Ϣ��<%
					end if
				%></td>
			</tr>
			<tr>
				<td  class=tablebody1 valign=top align=left><b>��Ϣ���⣺<%=htmlencode(rs("title"))%></b><hr size=1><%=dvbcode(rs("content"),4,3)%></td>
			</tr>
			<%rs.close
			set rs=nothing
			sql="select id,sender from message where incept='"&membername&"' and flag=0 and issend=1 and id>"&cstr(request("id")&" order by sendtime")
			set rs=conn.execute(sql)
			if not (rs.eof and rs.bof) then%>
				<tr>
					<td  class=tablebody2 valign=top align=right><a href=messanger.asp?action=read&id=<%=rs(0)%>&sender=<%=rs(1)%>>[��ȡ��һ����Ϣ]</a></td>
				</tr>
			<%end if
			rs.close
			set rs=nothing%>
		</table>
	<%end if
end sub
'ת����Ϣ
sub fw()
	stats="���Ͷ���"
	dim title,content,sender
	if request("id")<>"" and isNumeric(request("id")) then
		set rs=server.createobject("adodb.recordset")
		sql="select title,content,sender from message where (incept='"&membername&"' or sender='"&membername&"') and id="&request("id")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			Errmsg=Errmsg+"<br>"+"<li>��ѡ����ز�����"
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
	  	<th colspan=2 height=25><input type=hidden name="action" value="send">���Ͷ���Ϣ--����������������Ϣ</th>
		</tr>
		<tr> 
	  	<td class=tablebody1 valign=middle width=15%><b>�ռ��ˣ�</b></td>
	  	<td  class=tablebody1 valign=middle><input type=text name="touser" value="<%=request("touser")%>" size=50> <SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)><OPTION selected value="">ѡ��</OPTION><%
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
			<td class=tablebody1 valign=top><b>���⣺</b></td>
			<td class=tablebody1 valign=middle><input type=text name="title" size=70 maxlength=80 value="Fw��<%=title%>">&nbsp;</td>
		</tr>
		<tr> 
			<td class=tablebody1 valign=top><b>���ݣ�</b></td>
			<td class=tablebody1 valign=middle><textarea cols=70 rows=6 name="message" title="Ctrl+Enter����"><%
				response.write "========== ������ת����Ϣ ========="
				response.write vbNewLine
				response.write "ԭ�����ˣ�"&sender&vbNewLine&vbNewLine
				response.write server.htmlencode(content)
				response.write "==================================="
			%></textarea></td>
		</tr>
		<tr> 
			<td  class=tablebody1 colspan=2><b>˵��</b>��<br>�� ������ʹ��<b>Ctrl+Enter</b>����ݷ��Ͷ���<br>�� ������Ӣ��״̬�µĶ��Ž��û�������ʵ��Ⱥ�������<b><%=GroupSetting(33)%></b>���û�<br>�� �������<b>50</b>���ַ����������<b><%=GroupSetting(34)%></b>���ַ�<br><font color=<%=forum_body(8)%>>�� ÿ����һ���ռ�����ȡ����<%
				if isnull(myvip) or myvip<>1 then
					response.write forum_user(26)
				else
					response.write forum_user(27)
				end if
			%>Ԫ�����������ֽ�<%=mymoney%>Ԫ</font></td>
		</tr>
		<tr> 
			<td class=tablebody2 valign=middle colspan=2 align=center> <input type=Submit value="����" name=Submit>&nbsp; <input type=Submit value="����" name=Submit>&nbsp; <input type="reset" name="Clear" value="���">&nbsp; <input type="button" name="close" value="�ر�" onclick="window.close()"></td>
		</tr>
	</table>
	</form>
<%end sub

sub savemsg()
	stats="���Ͷ��ųɹ�"
	dim incept,title,message,subtype
	if request("touser")="" then
		errmsg=errmsg+"<br>"+"<li>��������д���Ͷ����˰ɡ�"
		founderr=true
		exit sub
	else
		incept=CheckStr(request("touser"))
		incept=split(incept,",")
	end if
	if request("title")="" then
		errmsg=errmsg+"<br>"+"<li>����û����д����ѽ��"
		founderr=true
		exit sub
	elseif strlength(request("title"))>50 then
		errmsg=errmsg+"<br>"+"<li>�����޶����50���ַ���"
		founderr=true
		exit sub
	else
		title=CheckStr(request("title"))
	end if
	if request("message")="" then
		errmsg=errmsg+"<br>"+"<li>�����Ǳ���Ҫ��д���ޡ�"
		founderr=true
		exit sub
	elseif strlength(request("message"))>Cint(GroupSetting(34)) then
		errmsg=errmsg+"<br>"+"<li>�����޶����"&GroupSetting(34)&"���ַ���"
		founderr=true
		exit sub
	else
		message=CheckStr(request("message"))
	end if
	if isnull(myvip) or myvip<>1 then
		if mymoney<(ubound(incept)+1)*clng(forum_user(26)) then
			errmsg=errmsg+"<br>"+"<li>�����ֽ𲻹�����Щ�˷��Ͷ��š�"
			founderr=true
			exit sub
		end if
	else
		if mymoney<(ubound(incept)+1)*clng(forum_user(27)) then
			errmsg=errmsg+"<br>"+"<li>�����ֽ𲻹�����Щ�˷��Ͷ��š�"
			founderr=true
			exit sub
		end if
	end if
	for i=0 to ubound(incept)
		sql="select username,msgflag from [user] where username='"&replace(incept(i),"'","")&"'"
		set rs=server.createobject("ADODB.Recordset")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			errmsg=errmsg+"<br>"+"<li>��̳û���û���"&replace(incept(i),"'","")&"����������ķ��Ͷ���д�����"
			founderr=true
			exit sub
		elseif not isnull(rs("msgFlag")) and rs("msgflag")=9 then
			errmsg=errmsg+"<br>"+"<li>�û���"&replace(incept(i),"'","")&"��Ŀǰ�������κζ���Ϣ��"
			founderr=true
			exit sub
		elseif not isnull(rs("msgFlag")) and rs("msgflag")=1 then
			if conn.execute("select top 1 f_friend from friend where f_username='"&replace(incept(i),"'","")&"' and f_friend='"&membername&"'").eof then
				errmsg=errmsg+"<br>"+"<li>�û���"&replace(incept(i),"'","")&"��Ŀǰֻ���պ��ѵĶ���Ϣ��"
				founderr=true
				exit sub
			end if
		end if
		dim rst
		set rst=conn.execute("select top 1 F_id from [Friend] where F_username='"&replace(incept(i),"'","")&"' and F_friend='"&membername&"' and F_type=1")
		if not (rst.eof and rst.bof) then
			errmsg=errmsg+"<br>"+"<li>���Ѿ���"&rs(0)&"�����ˣ������ܸ������Ͷ��ţ�"
			founderr=true
			exit sub
		end if
		rst.close
		set rs=nothing
		if request("Submit")="����" then
			sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,1)"
			subtype="�ѷ�����Ϣ"
			if isnull(myvip) or myvip<>1 then
				conn.execute("update [user] set userwealth=userwealth-"&forum_user(26)&" where username='"&membername&"'")
			else
				conn.execute("update [user] set userwealth=userwealth-"&forum_user(27)&" where username='"&membername&"'")
			end if
		elseif request("Submit")="����" then
			sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,0)"
			subtype="������"
		else
			sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,1)"
			subtype="�ѷ�����Ϣ"
			if isnull(myvip) or myvip<>1 then
				conn.execute("update [user] set userwealth=userwealth-"&forum_user(26)&" where username='"&membername&"'")
			else
				conn.execute("update [user] set userwealth=userwealth-"&forum_user(27)&" where username='"&membername&"'")
			end if
		end if
		conn.execute(sql)
		if i>Cint(GroupSetting(33))-1 then
			errmsg=errmsg+"<br>"+"<li>���ֻ�ܷ��͸�"&GroupSetting(33)&"���û�����������"&GroupSetting(33)&"λ�Ժ�������·���"
			founderr=true
			exit sub
			exit for
		end if
	next
	sucmsg=sucmsg+"<br>"+"<li><b>��ϲ�������Ͷ���Ϣ�ɹ���</b><br>���͵���Ϣͬʱ����������"&subtype&"�С�"
	call dvbbs_suc()
end sub

'������Ϣ
sub edit()
	stats="�޸Ķ���"
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
			Errmsg=Errmsg+"<br>"+"<li>û���ҵ���Ҫ�༭����Ϣ��"
			Founderr=true
			exit sub
		end if
		rs.close
		set rs=nothing
	else
		Errmsg=Errmsg+"<br>"+"<li>��ָ����ز�����"
		Founderr=true
		exit sub
	end if%>
	<form action="messanger.asp" method=post name=messager>
	<table cellpadding=3 cellspacing=1 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
		<tr> 
			<th colspan=2 height=25><input type=hidden name="action" value="savedit"><input type=hidden name="id" value="<%=id%>">���Ͷ���Ϣ--����������������Ϣ</th>
		</tr>
		<tr> 
			<td  class=tablebody1 valign=middle width=70><b>�ռ��ˣ�</b></td>
			<td  class=tablebody1 valign=middle><input type=text name="touser" value="<%=incept%>" size=70></td>
		</tr>
		<tr> 
			<td  class=tablebody1 valign=top><b>���⣺</b></td>
			<td  class=tablebody1 valign=middle><input type=text name="title" size=70 maxlength=80 value="<%=title%>"></td>
		</tr>
		<tr> 
			<td  class=tablebody1 valign=top><b>���ݣ�</b></td>
			<td  class=tablebody1 valign=middle><textarea cols=70 rows=8 name="message" title=""><%=server.htmlencode(content)%></textarea></td>
		</tr>
		<tr> 
			<td  class=tablebody1 colspan=2><b>˵��</b>��<br>�� ������ʹ��<b>Ctrl+Enter</b>����ݷ��Ͷ���<br>�� �������<b>50</b>���ַ����������<b><%=GroupSetting(34)%></b>���ַ�<BR><font color=<%=forum_body(8)%>>�� ÿ����һ���ռ�����ȡ����<%
				if isnull(myvip) or myvip<>1 then
					response.write forum_user(26)
				else
					response.write forum_user(27)
				end if
			%>Ԫ�����������ֽ�<%=mymoney%>Ԫ</font></td>
		</tr>
		<tr> 
			<td  class=tablebody2 valign=middle colspan=2 align=center><input type=Submit value="����" name=Submit>&nbsp; <input type=Submit value="����" name=Submit>&nbsp; <input type="reset" name="Clear" value="���">&nbsp; <input type="button" name="close" value="�ر�" onclick="window.close()"></td>
		</tr>
	</table>
	</form>
<%end sub

sub savedit()
	dim incept,title,message,subtype
	if request("id")="" or not isNumeric(request("id")) then
		Errmsg=Errmsg+"<br>"+"<li>��ָ����ز�����"
		Founderr=true
		exit sub
	end if
	if request("touser")="" then
		errmsg=errmsg+"<br>"+"<li>��������д���Ͷ����˰ɡ�"
		founderr=true
		exit sub
	else
		incept=checkStr(request("touser"))
	end if
	if request("title")="" then
		errmsg=errmsg+"<br>"+"<li>����û����д����ѽ��"
		founderr=true
		exit sub
	else
		title=checkStr(request("title"))
	end if
	if request("message")="" then
		errmsg=errmsg+"<br>"+"<li>�����Ǳ���Ҫ��д���ޡ�"
		founderr=true
		exit sub
	else
		message=checkStr(request("message"))
	end if
	sql="select username from [user] where username='"&incept&"'"
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		errmsg=errmsg+"<br>"+"<li>��̳û���û�"&incept&"��������ķ��Ͷ���д�����"
		founderr=true
		exit sub
	end if
	dim rst
	set rst=conn.execute("select top 1 F_id from [Friend] where F_username='"&incept&"' and F_friend='"&membername&"' and F_type=1")
	if not (rst.eof and rst.bof) then
		errmsg=errmsg+"<br>"+"<li>���Ѿ���"&rs(0)&"�����ˣ������ܸ������Ͷ��ţ�"
		founderr=true
		exit sub
	end if
	rst.close
	set rs=nothing
	if request("Submit")="����" then
		if isnull(myvip) or myvip<>1 then
			if mymoney<clng(forum_user(26)) then
				errmsg=errmsg+"<br>"+"<li>�����ֽ𲻹����������š�"
				founderr=true
				exit sub
			end if
		else
			if mymoney<clng(forum_user(27)) then
				errmsg=errmsg+"<br>"+"<li>�����ֽ𲻹����������š�"
				founderr=true
				exit sub
			end if
		end if
		sql="update message set incept='"&incept&"',sender='"&membername&"',title='"&title&"',content='"&message&"',sendtime=Now(),flag=0,issend=1 where id="&request("id")
		subtype="�ѷ�����Ϣ"
		if isnull(myvip) or myvip<>1 then
			conn.execute("update [user] set userwealth=userwealth-"&forum_user(26)&" where username='"&membername&"'")
		else
			conn.execute("update [user] set userwealth=userwealth-"&forum_user(27)&" where username='"&membername&"'")
		end if
	else
		sql="update message set incept='"&incept&"',sender='"&membername&"',title='"&title&"',content='"&message&"',sendtime=Now(),flag=0,issend=0 where id="&request("id")
		subtype="������"
	end if
	conn.execute(sql)
	sucmsg=sucmsg+"<br>"+"<li><b>��ϲ�������Ͷ���Ϣ�ɹ���</b><br>���͵���Ϣͬʱ����������"&subtype&"�С�"
	call dvbbs_suc()
end sub

'�ռ��߼�ɾ�������ڻ���վ������ֶ�delR������������������ɾ��
sub delinbox()
	dim delid
	delid=replace(request("id"),"'","")
	delid=replace(delid,";","")
	delid=replace(delid,"--","")
	delid=replace(delid,")","")
	if delid="" or isnull(delid) then
	Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
	Founderr=true
	else
		conn.execute("update message set delR=1 where incept='"&trim(membername)&"' and id in ("&delid&")")
		sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��"
		call dvbbs_suc()
	end if
end sub

sub AllDelinbox()
	conn.execute("update message set delR=1 where incept='"&trim(membername)&"' and delR=0")
	sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��"
	call dvbbs_suc()
end sub

'�����߼�ɾ�������ڻ���վ������ֶ�delS������������������ɾ��
sub deloutbox()
	dim delid
	delid=replace(request("id"),"'","")
	delid=replace(delid,";","")
	delid=replace(delid,"--","")
	delid=replace(delid,")","")
	if delid="" or isnull(delid) then
		Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
		Founderr=true
	else
		conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and issend=0 and id in ("&delid&")")
		sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��"
		call dvbbs_suc()
	end if
end sub

sub AllDeloutbox()
	conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and delS=0 and issend=0")
	sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��"
	call dvbbs_suc()
end sub

'�ѷ����߼�ɾ�������ڻ���վ������ֶ�delS������������������ɾ��
'delS��0δ������1������ɾ����2�����ߴӻ���վɾ��
sub delissend()
	dim delid
	delid=replace(request("id"),"'","")
	delid=replace(delid,";","")
	delid=replace(delid,"--","")
	delid=replace(delid,")","")
	if delid="" or isnull(delid) then
		Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
		Founderr=true
	else
		conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and issend=1 and id in ("&delid&")")
		sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��"
		call dvbbs_suc()
	end if
end sub

sub AllDelissend()
	conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and delS=0 and issend=1")
	sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��"
	call dvbbs_suc()
end sub

'�û�����ȫɾ���յ���Ϣ���߼�ɾ����������Ϣ���߼�ɾ����������Ϣ��������ֶ�delS����Ϊ2
sub delrecycle()
	dim delid
	delid=replace(request("id"),"'","")
	delid=replace(delid,";","")
	delid=replace(delid,"--","")
	delid=replace(delid,")","")
	if delid="" or isnull(delid) then
		Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
		Founderr=true
		exit sub
	else
		conn.execute("delete from message where incept='"&membername&"' and delR=1 and id in ("&delid&")")
		conn.execute("update message set delS=2 where sender='"&trim(membername)&"' and delS=1 and id in ("&delid&")")
		sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ�����ɻָ���"
		call dvbbs_suc()
	end if
end sub

sub AllDelrecycle()
	conn.execute("delete from message where incept='"&membername&"' and delR=1")	
	conn.execute("update message set delS=2 where sender='"&trim(membername)&"' and delS=1")
	sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ�����ɻָ���"
	call dvbbs_suc()
end sub

sub delete()
	dim delid
	delid=checkstr(request("id"))
	if not isNumeric(request("id")) or delid="" or isnull(delid) then
		Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
		Founderr=true
	else
		conn.execute("update message set delR=1 where incept='"&trim(membername)&"' and id="&delid)
		conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and id="&delid)
		sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ���������Ļ���վ�ڡ�"
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