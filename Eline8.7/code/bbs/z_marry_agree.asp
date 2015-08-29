<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_marry_conn.asp"-->
<%
dim SucFlag

SucFlag=false
stats="社区教堂"
call nav()
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入本社区教堂的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
	call dvbbs_error()
else
	call marry_nav()
	call main()
end if
call activeonline()
call footer()

sub main()
	dim agree,id,name,SendMsg
	founderr=false
	agree=request("agree")
	id=request("id")
	name=request("name")
	set rs=server.createobject("adodb.recordset")
	rs.open"select * from qiuhun where toUserName='"&membername&"'",connj,3,3
	if rs.bof then 
		rs.close
		Errmsg=Errmsg+"<br>"+"<li>你不是被求婚对象或者不正当使用了本社区的功能，请重新<a href=login.asp>登陆</a>或者同管理员联系。"
		call dvbbs_error()
	else
		rs.close
		if agree="yes" then
			set rs1 = Server.CreateObject("ADODB.recordset")
			rs1.open"Select * From [qiuhun] where id="&id&" and tousername='"&membername&"' and username='"&name&"'",connj,3,3
			if not rs1.eof then
				rs1("dlg")=true
				rs1.update
				message=rs1("message")
				rs.open"select * from [qiuhun] where tousername='"&rs1("tousername")&"' and dlg=false",connj,3,3
				ount1=0
				do while not rs.eof
					sql="insert into message(sender,incept,title,content,sendtime,issend)values('婚姻事务处','"&rs("username")&"','很遗憾的通知您','"&membername&"拒绝了"&rs("username")&"您的求婚，忘了吧，这种事常有的。','"&now()&"','1')"
					conn.execute sql
					rs.delete
					rs.movenext
					ount1=ount1+1
				loop
				rs.close
				rs.open"select * from [qiuhun] where username='"&membername&"'",connj,3,3
				if not rs.eof then
					name3=rs("tousername")
					sql="insert into message(sender,incept,title,content,sendtime,issend)values('婚姻事务处','"&name3&"','很遗憾的通知您','"&membername&"接受了"&name&"的求婚，所以"&membername&"对您的求婚已经无效。忘了吧，这种事常有的。','"&now()&"','1')"
					conn.execute sql
					rs.delete
				else
					name3=""
				end if
				rs.close
				rs1.close
				rs1.open"select * from [jie]",connj,3,3
				rs1.addnew
				rs1("username")=name
				rs1("thename")=membername
				rs1("content")=message
				rs1("year")=now()
				rs1("jiehun")=false
				rs1.update
				rs1.close
				sql="insert into message(sender,incept,title,content,sendtime,issend)values('婚姻事务处','"&name&"','恭喜！','"&membername&"答应了"&name&"您的求婚,你们商量个好日子然后去结婚教堂登记结婚吧','"&now()&"','1')"
				conn.execute sql
				if ount1>0 then
					SendMsg=SendMsg+"<br>"+"<li>系统已经发送了拒绝信息给另外"& ount1 &"位追求你的人。"
					founderr=true
					'Errmsg=Errmsg+"<br>"+"<li>系统已经发送了拒绝信息给另外"& ount1 &"位追求你的人。"
					'call dvbbs_error()
				end if
				if name3 <>"" then
					SendMsg=SendMsg+"<br>"+"<li>系统已经发送信息给"& name3 &"，遗憾的告诉他(她)您已经接受了"& name &"的求婚。"
					founderr=true
					'Errmsg=Errmsg+"<br>"+"<li>系统已经发送信息给"& name3 &"，遗憾的告诉他(她)您已经接受了"& name &"的求婚。"
					'call dvbbs_error()
				end if
				SendMsg=SendMsg+"<br>"+"<li>结婚成功！！系统已经发信息给"& name &"了，请你们商量好就去<a href=z_marry_register.asp>结婚登记</a>处登记吧。"
				founderr=true
				'Errmsg=Errmsg+"<br>"+"<li>结婚成功！！系统已经发信息给"& name &"了，请你们商量好就去<a href=z_marry_register.asp>结婚登记</a>处登记吧。"
				'call dvbbs_error()
				'if founderr=true
				SucFlag=true
				call agree_head()
				call agree_msg(SendMsg)
				'end if
			else
				rs1.close
				'SendMsg=SendMsg+"<br>"+"<li>不正当使用了本社区的功能，请重新<a href=login.asp>登陆</a>或者同管理员联系。"
				'founderr=true
				Errmsg=Errmsg+"<br>"+"<li>不正当使用了本社区的功能，请重新<a href=login.asp>登陆</a>或者同管理员联系。"
				call dvbbs_error()
			end if
		else
			connj.execute"delete from [qiuhun]  where id="&id 
			sql="insert into message(sender,incept,title,content,sendtime,issend)values('婚姻事务处','"&name&"','很遗憾的通知您：','"&membername&"拒绝了"&name&"您的求婚，忘了吧，这种事常有的。','"&now()&"','1')"
			conn.execute sql
				SendMsg=SendMsg+"<br>"+"<li>系统已经发送拒绝信息给"& name &"了，多情自古空余恨。"
				founderr=true
			SucFlag=true
			call agree_head()
			call agree_msg(SendMsg)
			'Errmsg=Errmsg+"<br>"+"<li>系统已经发送拒绝信息给"& name &"了，多情自古空余恨。"
			'call dvbbs_error()
		end if 
	end if
end sub

sub agree_head()%>
	<html><head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title><%=Forum_info(0)%>--求婚信息--</title>
	<!--#include file=inc/forum_css.asp-->
	</head>
	<body <%=Forum_body(11)%>" onkeydown="if(event.keyCode==13 && event.ctrlKey && document.messager=='[object]')messager.submit()">
<%end sub

sub agree_msg(strMsg)%>
	<br>
	<table cellpadding=3 cellspacing=1 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
		<tr align=center>
			<th width="100%">求婚信息</th>
		</tr>
		<tr>
			<td width="100%" class=tablebody1><b>操作结果：</b><br><%
				if founderr then
					response.write "<font color="&forum_body(8)&">"
					response.write strMsg
					response.write "</font>"
				else
					response.write strMsg
				end if
			%></td>
		</tr>
		<tr align=center>
			<td width="100%" class=tablebody2><a href="javascript:window.location.href='z_marry_manager.asp'">[返回月老版块]</a>&nbsp;&nbsp;<%
				if SucFlag then
					%><a href="z_visual_cartbag.asp">[返回购物袋]</a><%
				'else
					%><!--<a href="javascript:history.go(-1)">[返回上页]</a>--><%
				end if
			%></td>
		</tr>
	</table>
<%end sub%>