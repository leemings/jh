<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_marry_conn.asp"-->
<%
dim SucFlag

SucFlag=false
stats="��������"
call nav()
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û�н��뱾�������õ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
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
		Errmsg=Errmsg+"<br>"+"<li>�㲻�Ǳ���������߲�����ʹ���˱������Ĺ��ܣ�������<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
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
					sql="insert into message(sender,incept,title,content,sendtime,issend)values('��������','"&rs("username")&"','���ź���֪ͨ��','"&membername&"�ܾ���"&rs("username")&"������飬���˰ɣ������³��еġ�','"&now()&"','1')"
					conn.execute sql
					rs.delete
					rs.movenext
					ount1=ount1+1
				loop
				rs.close
				rs.open"select * from [qiuhun] where username='"&membername&"'",connj,3,3
				if not rs.eof then
					name3=rs("tousername")
					sql="insert into message(sender,incept,title,content,sendtime,issend)values('��������','"&name3&"','���ź���֪ͨ��','"&membername&"������"&name&"����飬����"&membername&"����������Ѿ���Ч�����˰ɣ������³��еġ�','"&now()&"','1')"
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
				sql="insert into message(sender,incept,title,content,sendtime,issend)values('��������','"&name&"','��ϲ��','"&membername&"��Ӧ��"&name&"�������,����������������Ȼ��ȥ�����õǼǽ���','"&now()&"','1')"
				conn.execute sql
				if ount1>0 then
					SendMsg=SendMsg+"<br>"+"<li>ϵͳ�Ѿ������˾ܾ���Ϣ������"& ount1 &"λ׷������ˡ�"
					founderr=true
					'Errmsg=Errmsg+"<br>"+"<li>ϵͳ�Ѿ������˾ܾ���Ϣ������"& ount1 &"λ׷������ˡ�"
					'call dvbbs_error()
				end if
				if name3 <>"" then
					SendMsg=SendMsg+"<br>"+"<li>ϵͳ�Ѿ�������Ϣ��"& name3 &"���ź��ĸ�����(��)���Ѿ�������"& name &"����顣"
					founderr=true
					'Errmsg=Errmsg+"<br>"+"<li>ϵͳ�Ѿ�������Ϣ��"& name3 &"���ź��ĸ�����(��)���Ѿ�������"& name &"����顣"
					'call dvbbs_error()
				end if
				SendMsg=SendMsg+"<br>"+"<li>���ɹ�����ϵͳ�Ѿ�����Ϣ��"& name &"�ˣ������������þ�ȥ<a href=z_marry_register.asp>���Ǽ�</a>���Ǽǰɡ�"
				founderr=true
				'Errmsg=Errmsg+"<br>"+"<li>���ɹ�����ϵͳ�Ѿ�����Ϣ��"& name &"�ˣ������������þ�ȥ<a href=z_marry_register.asp>���Ǽ�</a>���Ǽǰɡ�"
				'call dvbbs_error()
				'if founderr=true
				SucFlag=true
				call agree_head()
				call agree_msg(SendMsg)
				'end if
			else
				rs1.close
				'SendMsg=SendMsg+"<br>"+"<li>������ʹ���˱������Ĺ��ܣ�������<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
				'founderr=true
				Errmsg=Errmsg+"<br>"+"<li>������ʹ���˱������Ĺ��ܣ�������<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
				call dvbbs_error()
			end if
		else
			connj.execute"delete from [qiuhun]  where id="&id 
			sql="insert into message(sender,incept,title,content,sendtime,issend)values('��������','"&name&"','���ź���֪ͨ����','"&membername&"�ܾ���"&name&"������飬���˰ɣ������³��еġ�','"&now()&"','1')"
			conn.execute sql
				SendMsg=SendMsg+"<br>"+"<li>ϵͳ�Ѿ����;ܾ���Ϣ��"& name &"�ˣ������Թſ���ޡ�"
				founderr=true
			SucFlag=true
			call agree_head()
			call agree_msg(SendMsg)
			'Errmsg=Errmsg+"<br>"+"<li>ϵͳ�Ѿ����;ܾ���Ϣ��"& name &"�ˣ������Թſ���ޡ�"
			'call dvbbs_error()
		end if 
	end if
end sub

sub agree_head()%>
	<html><head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title><%=Forum_info(0)%>--�����Ϣ--</title>
	<!--#include file=inc/forum_css.asp-->
	</head>
	<body <%=Forum_body(11)%>" onkeydown="if(event.keyCode==13 && event.ctrlKey && document.messager=='[object]')messager.submit()">
<%end sub

sub agree_msg(strMsg)%>
	<br>
	<table cellpadding=3 cellspacing=1 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
		<tr align=center>
			<th width="100%">�����Ϣ</th>
		</tr>
		<tr>
			<td width="100%" class=tablebody1><b>���������</b><br><%
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
			<td width="100%" class=tablebody2><a href="javascript:window.location.href='z_marry_manager.asp'">[�������ϰ��]</a>&nbsp;&nbsp;<%
				if SucFlag then
					%><a href="z_visual_cartbag.asp">[���ع����]</a><%
				'else
					%><!--<a href="javascript:history.go(-1)">[������ҳ]</a>--><%
				end if
			%></td>
		</tr>
	</table>
<%end sub%>