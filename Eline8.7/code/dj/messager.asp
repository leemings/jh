<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<!--#include file="function.asp"-->
<!--#include file="inc/char.inc"-->
<html>
<head>
<title><%=Websitename%>--����Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
td
{
FONT-WEIGHT: normal; FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; color: #FFFFFF
}
A {
	TEXT-DECORATION: none
}
A:link {
	COLOR: #FFFFFF
}
A:visited {
	COLOR: #FFFFFF
}
A:active {
	COLOR: #FFFFFF
}
A:hover {
	COLOR: #FFFFFF; TEXT-DECORATION: underline overline
}
input
{
		background-image: url('images/inbg.gif');border:1px solid #CE9A00; padding-left: 0
}

</style>
<script src=js/js.js></script>
</head>
<%
dim msg
if request.cookies("username")="" then
	conn.close
	set conn=nothing
			errmsg="<li>������½���ٽ��иò�����</li>"
			call error()
			response.end
end if
set rs=server.createobject("adodb.recordset")

if founderr=true then
	call error()
else
	action=request("action")
	select case action
	case "inbox"
		call inbox()
	case "outbox"
		call outbox()
	case "new"
		call sendmsg()
	case "read"
		call read()
	case "outread"
		call read()
	case "delete"
		call delete()
	case "deleteall"
		call deleteall()
	case "send"
		call savemsg()
	case "newmsg"
		call newmsg()
	case "friendlist"
		call friendlist()
	case "friendadd"
		call friendadd()
	case "frienddel"
		call FriendDel()
	case else
		conn.close
		set conn=nothing
			errmsg="<li>���������������������������Ա��ϵ!!!</li>"
			call error()
			response.end
	end select
end if
'�ռ���
sub inbox()
%>
<body topmargin=0 leftmargin=0 onkeydown="if(event.keyCode==13 && event.ctrlKey)messager.submit()" text="#000000">
<div align="center">
  <center>
    <table cellpadding=0 cellspacing=0 border=0 width="100%" style="border-collapse: collapse">
      <tr>
        <td width="100%">
          <div align="center">
            <center>
          <table cellpadding=0 cellspacing=0 border=1 width=400 bordercolor="#000000" style="border-collapse: collapse">
            <tr> 
              <td align=center colspan=3 width="398" bgcolor="#F09F46"><b>��ӭʹ�������ռ��䣬<%=request.cookies("username")%></b></td>
            </tr>
            <tr>
              <td align=center colspan=3 bgcolor="#C68600" width="398"><a href="messager.asp?action=inbox"><img  src="images/inboxpm.gif" border=0 alt="�ռ���"></a> &nbsp; <a href="messager.asp?action=outbox">
                 <img  src="images/outboxpm.gif" border=0 alt="������"></a>&nbsp;&nbsp; <a href="messager.asp?action=friendlist"><img border="0" src="images/friendlist.gif"></a> &nbsp; <a href="messager.asp?action=new"><img  src="images/newpm.gif" border=0 alt="������Ϣ"></a></td>
            </tr>
            <tr>
              <td align=center width="77" bgcolor="#F09F46"><b>����</b></td>
              <td align=center width="190" bgcolor="#F09F46"><b>����</b></td>
              <td align=center width="129" bgcolor="#F09F46"><b>�Ѷ�</b></td>
            </tr>
<%
	
	sql="select * from message where incept='"&request.cookies("username")&"' order by flag"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
            <tr>
              <td align=center colspan=3 bgcolor="#C68600" width="398">����û���������ޣ���</td>
            </tr>
<%	else
		do while not rs.eof%>
            <tr>
              <td align=center bgcolor="#C68600" width="77"><%=rs("sender")%>��</td>
              <td align=left bgcolor="#C68600" width="190"><a href="messager.asp?action=read&id=<%=rs("id")%>"><%=htmlencode1(rs("title"))%></a>��</td>
              <td align=center bgcolor="#C68600" width="129"><%if rs("flag")=0 then%><font color=red><b>��</b></font><%else%>��<%end if%></td>
            </tr>
<%
			rs.movenext
		loop
	end if
	rs.close
%>
                
            <tr> 
              <td align=center colspan=3 bgcolor="#F09F46" width="398"><a href="messager.asp?action=deleteall" >ɾ�����еĶ���Ϣ</a></td>
            </tr>
          </table>
            </center>
          </div>
        </td>
      </tr>
    </table>
  </center>
</div>
<%
end sub
'������
sub outbox()
%>
<body topmargin=0 leftmargin=0 onkeydown="if(event.keyCode==13 && event.ctrlKey)messager.submit()" text="#000000">
<div align="center">
  <center>
    <table cellpadding=0 cellspacing=0 border=0 width="100%" style="border-collapse: collapse">
      <tr>
        <td>
          <div align="center">
            <center>
          <table cellpadding=0 cellspacing=0 border=1 width=400 bordercolor="#000000" style="border-collapse: collapse">
            <tr>
              <td align=center colspan=2 bgcolor="#F09F46"><b>��ӭʹ�ö���Ϣ���ͣ�<%=request.cookies("username")%></b></td>
            </tr>
            <tr>
              <td align=center colspan=3 bgcolor="#C68600"><a href="messager.asp?action=inbox"><img  src="images/inboxpm.gif" border=0 alt="�ռ���"></a> &nbsp; <a href="messager.asp?action=outbox">
                 <img  src="images/outboxpm.gif" border=0 alt="������"></a>&nbsp;&nbsp; <a href="messager.asp?action=friendlist"><img border="0" src="images/friendlist.gif"></a> &nbsp; <a href="messager.asp?action=new"><img  src="images/newpm.gif" border=0 alt="������Ϣ"></a></td>
            </tr>
            <tr>
              <td align=center width=20% bgcolor="#F09F46"><b>�ռ���</b></td>
              <td align=center bgcolor="#F09F46"><b>����</b></td>
            </tr>
<%
	sql="select * from message where sender='"&request.cookies("username")&"'"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
            <tr>
              <td align=center colspan=2 bgcolor="#C68600">����û�и����˷�����Ϣ��~~</td>
            </tr>
<%	else
		do while not rs.eof%>
            <tr>
              <td align=center width=20% bgcolor="#C68600"><%=rs("incept")%>��</td>
              <td align=left bgcolor="#C68600"><a href="messager.asp?action=outread&id=<%=rs("id")%>"><%=htmlencode1(rs("title"))%></a>��</td>
            </tr>
<%
			rs.movenext
		loop
	end if
	rs.close
%>                
            <tr>
              <td align=center colspan=2 bgcolor="#F09F46"><font face="����" color=#333333><a href="messager.asp?action=deleteall">ɾ�����еĶ���Ϣ</a></font></td>
            </tr>
          </table>
            </center></div>
        </td>
      </tr>
    </table>
  </center>
</div>
<%
end sub
'������Ϣ
sub sendmsg()
%>
<body topmargin=0 leftmargin=0 onkeydown="if(event.keyCode==13 && event.ctrlKey)messager.submit()" text="#000000">
<form action="messager.asp" method=post name=messager>
  <div align="center">
    <center>
  <table cellpadding=0 cellspacing=0 border=0 width="100%" style="border-collapse: collapse">
    <tr> 
      <td> 
        <div align="center">
          <center>
        <table cellpadding=0 cellspacing=0 border=1 width=400 bordercolor="#000000" style="border-collapse: collapse">
          <tr> 
            <td align=center colspan=3 bgcolor="#F09F46"><b>���Ͷ���Ϣ</b></td>
          </tr>
          <tr> 
            <td align=center colspan=3 bgcolor="#C68600"><a href="messager.asp?action=inbox"><img  src="images/inboxpm.gif" border=0 alt="�ռ���"></a> 
              &nbsp; <a href="messager.asp?action=outbox"><img  src="images/outboxpm.gif" border=0 alt="������"></a>&nbsp;&nbsp; <a href="messager.asp?action=friendlist"><img border="0" src="images/friendlist.gif"></a> 
              &nbsp; <a href="messager.asp?action=new"><img  src="images/newpm.gif" border=0 alt="������Ϣ"></a></td>
          </tr>
          <tr> 
            <td colspan=2 align=center bgcolor="#F09F46"><input type=hidden name="action" value="send"><b>����������������Ϣ</b></td>
          </tr>
          <tr> 
            <td valign=middle bgcolor="#C68600"><b>�ռ��ˣ�</b></td>
            <td valign=middle bgcolor="#C68600">
            <input type=text name="touser" value="<%=request("touser")%>" size=20>
              <SELECT name=font onchange=document.messager.touser.value=options[selectedIndex].value>
              <OPTION selected value="">ѡ��</OPTION>
<%
set rs=server.createobject("adodb.recordset")
sql="select F_friend from Friend where F_username='"&request.cookies("username")&"' order by F_addtime desc"
rs.open sql,conn,1,1
do while not rs.eof
%>
			  <OPTION value="<%=rs(0)%>"><%=rs(0)%></OPTION> 
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
			  </SELECT>
            </td>
          </tr>
          <tr> 
            <td valign=top width=30% bgcolor="#C68600"><b>���⣺</b></td>
            <td valign=middle bgcolor="#C68600"><input type=text name="title" size=36 maxlength=80></td>
          </tr>
          <tr> 
            <td valign=top width=30% bgcolor="#C68600"><b>���ݣ�</b><br>Ctrl+Enter����</td>
            <td valign=middle bgcolor="#C68600"><textarea cols=35 rows=6 name="message" title="Ctrl+Enter����"></textarea></td>
          </tr>
          <tr> 
            <td colspan=2 align=center bgcolor="#F09F46"> 
              <input type=Submit value="�� ��" name=Submit&quot;>&nbsp;<input type="reset" name="Clear" value="�� ��">
            </td>
          </tr>
        </table>
          </center></div>
      </td>
    </tr>
  </table>
    </center>
  </div>
</form>
<%
end sub
'��ȡ��Ϣ
sub read()
%>
<%
	sql="update message set flag=1 where ID="&cstr(request("id"))
	rs.open sql,conn,1,3
	sql="select * from message where incept='"&request.cookies("username")&"' and id="&cstr(request("id"))
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		rs.Close
		set rs=nothing
		conn.close
		set conn=nothing
			errmsg="<li>�Բ������ǲ����ܱ��˵��������ˣ����������ˢ��!!!</li>"
			call error()
			response.end 
	else
%>
<body topmargin=0 leftmargin=0 onkeydown="if(event.keyCode==13 && event.ctrlKey)messager.submit()" text="#000000">
<div align="center">
  <center>
<table cellpadding=0 cellspacing=0 border=0 width="100%" style="border-collapse: collapse">
  <tr>
    <td>
      <div align="center">
        <center>
      <table cellpadding=0 cellspacing=0 border=1 width=400 bordercolor="#000000" style="border-collapse: collapse" bgcolor="#C68600">
        <tr>
          <td align=center colspan=3><b>��ӭʹ�ö���Ϣ���գ�<%=request.cookies("username")%></b></td>
        </tr>
        <tr>
          <td align=center colspan=3 bgcolor="#C68600"><a href="messager.asp?action=delete&id=<%=rs("id")%>"><img  src="images/deletepm.gif" border=0 alt="ɾ����Ϣ"></a> &nbsp; <a href="messager.asp?action=inbox">
             <img  src="images/inboxpm.gif" border=0 alt="�ռ���"></a> &nbsp; <a href="messager.asp?action=outbox"><img  src="images/outboxpm.gif" border=0 alt="������"></a>&nbsp;&nbsp; <a href="messager.asp?action=friendlist"><img border="0" src="images/friendlist.gif"></a>&nbsp; &nbsp;<a href="messager.asp?action=new"><img   src="images/newpm.gif" border=0 alt="������Ϣ"></a>&nbsp; &nbsp;<a href="messager.asp?action=new&touser=<%=rs("sender")%>"><img  src="images/replypm.gif" border=0 alt="�ظ���Ϣ"></a></td>
        </tr>
        <tr>
          <td valign=middle align=center>
<%if request("action")="outread" then%>
            ��<b><%=rs("sendtime")%></b>�������ʹ���Ϣ��<b><%=rs("incept")%></b>��
<%else%>
		    ��<b><%=rs("sendtime")%></b>��<b><%=rs("sender")%></b>�������͵���Ϣ��
<%end if%></font></td>
        </tr>
        <tr>
          <td valign=top align=left><b>��Ϣ���⣺<%=htmlencode1(rs("title"))%></b><p><%if rs("sender")="���ž���" then%><%=rs("content")%><%else%><%=htmlencode2(rs("content"))%><%end if%></td>
        </tr>
      </table>
        </center></div>
    </td>
  </tr>
</table>
  </center>
</div>
<%end if
end sub

sub savemsg()
	if request("touser")="" then
		conn.close
		set conn=nothing
			response.write "<script language=javascript>window.alert('��Ҫ����˭ѽ����ȥд�������ˣ�');history.back(1);</script>"
			response.end  
	elseif request("touser")="���ž���" then
			conn.close
		set conn=nothing
			response.write "<script language=javascript>window.alert('���ž�����ϵͳ�����ʺţ��㲻�ܷ�������������������ϵվ����');history.back(1);</script>"
			response.end 
	else
		incept=request("touser")
	end if
	if request("title")="" then
		conn.close
		set conn=nothing
			response.write "<script language=javascript>window.alert('����һ��Ҫ��д����ȥ���ϱ����ٷ��Ͱɣ�');history.back(1);</script>"
			response.end
	else
		title=request("title")
	end if
	if request("message")="" then
		conn.close
		set conn=nothing
			response.write "<script language=javascript>window.alert('����Ϣ�����ݲ���Ϊ��ѽ��д�����仰�ɣ�');history.back(1);</script>"
			response.end 
	else
		message=request("message")
	end if
	if founderr=true then
		call error()
	else
		sql="select * from [user] where username='"&incept&"'"
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			rs.Close 
			set rs=nothing
			conn.close
			set conn=nothing
			response.write "<script language=javascript>window.alert('�����˲��Ǳ�վ�û���');history.back(1);</script>"
			response.end 
		else
			rs.Close 
			sql="select * from message"
			rs.open sql,conn,1,3
			rs.addnew
			rs("incept")=incept
			rs("sender")=request.cookies("username")
			rs("title")=title
			rs("content")=message
			rs("sendtime")=now()
			rs("flag")=0
			rs.update
			msg=msg+"<br>"+"<li><b>��ϲ�������Ͷ���Ϣ�ɹ���</b><br>���͵���Ϣͬʱ���������ķ����䡣"
			call success()
			rs.close
		end if
	end if
end sub

sub delete()
	sql="delete from message where incept='"&request.cookies("username")&"' and id="&cstr(request("id"))
	conn.execute sql
	if err.Number<>0 then
		err.clear
		conn.close
		set conn=nothing
			response.write "<script language=javascript>window.alert('����Ϣɾ��ʧ�ܣ�');history.back(1);</script>"
			response.end 
	else
		msg=msg+"<br>"+"<li>����Ϣ�ɹ�ɾ����"
		call success()
	end if
end sub

sub deleteall()
	sql="delete from message where incept='"&request.cookies("username")&"'"
	conn.execute sql
	if err.Number<>0 then
		err.clear
		conn.close
		set conn=nothing
			response.write "<script language=javascript>window.alert('����Ϣɾ��ʧ�ܣ�');history.back(1);</script>"
			response.end 
	else
		msg=msg+"<br>"+"<li>����Ϣȫ���ɹ�ɾ����"
		call success()
	end if
end sub

sub friendadd
touser=checkStr(request("touser"))
REM �ж��û��Ƿ����
sql="select username from [user] where username='"&touser&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
			response.write "<script language=javascript>window.alert('�ܱ�Ǹ����վû������û���');history.back(1);</script>"
			response.end 
	exit sub
end if
if touser=request.cookies("username") then
			response.write "<script language=javascript>window.alert('����������Լ�Ϊ���ѣ�');history.back(1);</script>"
			response.end 
	exit sub
end if
if request("touser")="���ž���" then
			conn.close
		set conn=nothing
			response.write "<script language=javascript>window.alert('���ž�����ϵͳ�����ʺţ��㲻�ܼ���Ϊ���ѣ�����������ϵվ����');history.back(1);</script>"
			response.end
end if
sql="select F_friend from friend where F_username='"&request.cookies("username")&"' and  F_friend='"&touser&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	sql="insert into friend (F_username,F_friend,F_addtime) values ('"&request.cookies("username")&"','"&touser&"',Now())"
	conn.execute(sql)
end if
Response.Redirect"messager.asp?action=friendlist"
end sub
'ɾ������
sub frienddel
delid=request("id")
if delid="" or isnull(delid) then
			response.write "<script language=javascript>window.alert('��Ĳ�����������������ش���');history.back(1);</script>"
			response.end 
exit sub
else
	conn.execute("delete from Friend where F_username='"&request.cookies("username")&"' and F_id in ("&delid&")")
	Response.Redirect"messager.asp?action=friendlist"
end if
end sub

sub success()
%>
<body topmargin=0 leftmargin=0 onkeydown="if(event.keyCode==13 && event.ctrlKey)messager.submit()" text="#000000">
    <div align="center">
      <center>
    <table cellpadding=0 cellspacing=0 border=0 width="100%" style="border-collapse: collapse">
      <tr>
        <td>
          <div align="center">
            <center>
          <table cellspacing=0 border=1 width=400 bordercolor="#000000" style="border-collapse: collapse" cellpadding="0" bgcolor="#C68600">
            <tr align="center"> 
              <td width="100%" bgcolor="#" >�ɹ�������Ϣ</td>
            </tr>
            <tr> 
              <td width="100%"><%=msg%>��</td>
            </tr>
            <tr align="center"> 
              <td width="100%"><a href="javascript:history.go(-1)"> << ������һҳ</a> </td>
            </tr>  
          </table>
            </center></div>
        </td>
      </tr>
    </table>
      </center>
    </div>
<%
end sub
'����Ϣ��ʾ
sub newmsg()
%>
<body topmargin=0 leftmargin=0 onkeydown="if(event.keyCode==13 && event.ctrlKey)messager.submit()" text="#000000">
    <bgsound src="newmessage.mp3" loop="-1">
    <div align="center">
      <center>
    <table cellpadding=0 cellspacing=0 border=0 width="100%" style="border-collapse: collapse">
      <tr>
        <td>
          <div align="center">
            <center>
          <table cellspacing=0 border=1 width=400 bordercolor="#000000" style="border-collapse: collapse" cellpadding="0">
            <tr align="center"> 
              <td width="100%"  bgcolor="#F09F46">����Ϣ֪ͨ</td>
            </tr>
            <tr> 
              <td width="100%" align=center bgcolor="#C68600"><br>
                <a href=messager.asp?action=inbox><img src="images/newmail.gif" border=0>���µĶ���Ϣ</a><br>
                <br>
                <a href=messager.asp?action=inbox>���˲鿴</a><br><br>
              </td>
            </tr>
          </table>
            </center>
          </div>
        </td>
      </tr>
    </table>
      </center>
    </div>
<%
end sub
'�����б�
sub friendlist()
%>
<body topmargin=0 leftmargin=0 onkeydown="if(event.keyCode==13 && event.ctrlKey)messager.submit()" text="#000000">
<div align="center">
  <center>
    <table cellpadding=0 cellspacing=0 border=0 width="100%" style="border-collapse: collapse">
      <tr>
        <td>
          <div align="center">
            <center>
          <table cellpadding=0 cellspacing=0 border=1 width=400 bordercolor="#000000" style="border-collapse: collapse">
            <tr>
              <td align=center width="398" bgcolor="#F09F46"><b>�û�<%=request.cookies("username")%>�ĺ����б�</b></td>
            </tr>
            <tr>
              <td align=center bgcolor="#C68600" width="398"><a href="messager.asp?action=inbox"><img  src="images/inboxpm.gif" border=0 alt="�ռ���"></a> &nbsp; <a href="messager.asp?action=outbox">
                 <img  src="images/outboxpm.gif" border=0 alt="������"></a>&nbsp;&nbsp; <a href="messager.asp?action=friendlist"><img border="0" src="images/friendlist.gif"></a> &nbsp; <a href="messager.asp?action=new"><img  src="images/newpm.gif" border=0 alt="������Ϣ"></a></td>
            </tr>
            <tr>
              <td align=center bgcolor="#C68600" width="100%">

<div align="center">
  <center>
<table cellpadding=0 cellspacing=0 width="398" border="1" bordercolor="#000000" style="border-collapse: collapse">
            <tr>
                <td width="114" align="center">����</td> 
                <td width="121" align="center">�ʼ�</td> 
                <td width="92" align="center">OICQ</td>
                <td width="33" align="center">����</td> 
                <td width="32" align="center">ɾ��</td></tr>
<%
	sql="select F.*,U.id,U.email,U.homepage,U.oicq from Friend F inner join [user] U on F.F_Friend=U.username where F.F_username='"&request.cookies("username")&"' order by F.f_addtime desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
                <tr>
                <td class=tablebody1 align=center valign=middle colspan=5 width="396">���ĺ����б���û���κ����ݡ�</td>
                </tr>
<%else%>
<%do while not rs.eof%>
                <tr>
                    <td align=center valign=middle class=tablebody1 width="114"><a href=### onclick="javascript:open_window('UserInfo.asp?id=<%=rs("id")%>','UserInfo','scrollbars=no,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,copyhistory=no,top=10,left=10,width=435,height=448')"><%=htmlencode(rs("F_friend"))%></a>��</td>
                    <td align=center valign=middle class=tablebody1 width="121"><a href="mailto:<%=htmlencode(rs("email"))%>"><%=htmlencode(rs("email"))%></a>��</td>
                    <td align=center class=tablebody1 width="92"><%=htmlencode(rs("oicq"))%>��</td>
                    <td align=center class=tablebody1 width="33"><a href="messager.asp?action=new&touser=<%=htmlencode(rs("f_friend"))%>">����</a></td>
                <td align=center class=tablebody1 width="32">
                <a href="messager.asp?action=frienddel&id=<%=rs("f_id")%>">ɾ��</a></td>
                </tr>
<%
	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
%>
                </table>


  </center>
</div>


</td>
            </tr>
<form action="messager.asp" method=post name=messager2>
<input type=hidden name="action" value="friendadd">
            <tr>
              <td align=center bgcolor="#F09F46" width="398">
              <input type=text name="touser" value="������Ҫ���Ϊ���ѵ��û���" size=28>
              <input type=submit value="��  ��" name=Submit></td>
            </tr>
</form>
          </table>
            </center></div>
        </td>
      </tr>
    </table>
  </center>
</div>

<%
end sub
set rs=nothing
conn.close
set conn=nothing
%>
</body>
</html>