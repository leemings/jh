<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<!--#include file="function.asp"-->
<!--#include file="inc/char.inc"-->
<html>
<head>
<title><%=Websitename%>--短消息</title>
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
			errmsg="<li>请您登陆后再进行该操作。</li>"
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
			errmsg="<li>服务器发生错误，请与服务器管理员联系!!!</li>"
			call error()
			response.end
	end select
end if
'收件箱
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
              <td align=center colspan=3 width="398" bgcolor="#F09F46"><b>欢迎使用您的收件箱，<%=request.cookies("username")%></b></td>
            </tr>
            <tr>
              <td align=center colspan=3 bgcolor="#C68600" width="398"><a href="messager.asp?action=inbox"><img  src="images/inboxpm.gif" border=0 alt="收件箱"></a> &nbsp; <a href="messager.asp?action=outbox">
                 <img  src="images/outboxpm.gif" border=0 alt="发件箱"></a>&nbsp;&nbsp; <a href="messager.asp?action=friendlist"><img border="0" src="images/friendlist.gif"></a> &nbsp; <a href="messager.asp?action=new"><img  src="images/newpm.gif" border=0 alt="发送消息"></a></td>
            </tr>
            <tr>
              <td align=center width="77" bgcolor="#F09F46"><b>来自</b></td>
              <td align=center width="190" bgcolor="#F09F46"><b>主题</b></td>
              <td align=center width="129" bgcolor="#F09F46"><b>已读</b></td>
            </tr>
<%
	
	sql="select * from message where incept='"&request.cookies("username")&"' order by flag"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
            <tr>
              <td align=center colspan=3 bgcolor="#C68600" width="398">您还没有新留言噢：）</td>
            </tr>
<%	else
		do while not rs.eof%>
            <tr>
              <td align=center bgcolor="#C68600" width="77"><%=rs("sender")%>　</td>
              <td align=left bgcolor="#C68600" width="190"><a href="messager.asp?action=read&id=<%=rs("id")%>"><%=htmlencode1(rs("title"))%></a>　</td>
              <td align=center bgcolor="#C68600" width="129"><%if rs("flag")=0 then%><font color=red><b>否</b></font><%else%>是<%end if%></td>
            </tr>
<%
			rs.movenext
		loop
	end if
	rs.close
%>
                
            <tr> 
              <td align=center colspan=3 bgcolor="#F09F46" width="398"><a href="messager.asp?action=deleteall" >删除所有的短消息</a></td>
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
'发件箱
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
              <td align=center colspan=2 bgcolor="#F09F46"><b>欢迎使用短消息发送，<%=request.cookies("username")%></b></td>
            </tr>
            <tr>
              <td align=center colspan=3 bgcolor="#C68600"><a href="messager.asp?action=inbox"><img  src="images/inboxpm.gif" border=0 alt="收件箱"></a> &nbsp; <a href="messager.asp?action=outbox">
                 <img  src="images/outboxpm.gif" border=0 alt="发件箱"></a>&nbsp;&nbsp; <a href="messager.asp?action=friendlist"><img border="0" src="images/friendlist.gif"></a> &nbsp; <a href="messager.asp?action=new"><img  src="images/newpm.gif" border=0 alt="发送消息"></a></td>
            </tr>
            <tr>
              <td align=center width=20% bgcolor="#F09F46"><b>收件人</b></td>
              <td align=center bgcolor="#F09F46"><b>标题</b></td>
            </tr>
<%
	sql="select * from message where sender='"&request.cookies("username")&"'"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
            <tr>
              <td align=center colspan=2 bgcolor="#C68600">您还没有给别人发过信息呢~~</td>
            </tr>
<%	else
		do while not rs.eof%>
            <tr>
              <td align=center width=20% bgcolor="#C68600"><%=rs("incept")%>　</td>
              <td align=left bgcolor="#C68600"><a href="messager.asp?action=outread&id=<%=rs("id")%>"><%=htmlencode1(rs("title"))%></a>　</td>
            </tr>
<%
			rs.movenext
		loop
	end if
	rs.close
%>                
            <tr>
              <td align=center colspan=2 bgcolor="#F09F46"><font face="宋体" color=#333333><a href="messager.asp?action=deleteall">删除所有的短消息</a></font></td>
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
'发送信息
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
            <td align=center colspan=3 bgcolor="#F09F46"><b>发送短消息</b></td>
          </tr>
          <tr> 
            <td align=center colspan=3 bgcolor="#C68600"><a href="messager.asp?action=inbox"><img  src="images/inboxpm.gif" border=0 alt="收件箱"></a> 
              &nbsp; <a href="messager.asp?action=outbox"><img  src="images/outboxpm.gif" border=0 alt="发件箱"></a>&nbsp;&nbsp; <a href="messager.asp?action=friendlist"><img border="0" src="images/friendlist.gif"></a> 
              &nbsp; <a href="messager.asp?action=new"><img  src="images/newpm.gif" border=0 alt="发送消息"></a></td>
          </tr>
          <tr> 
            <td colspan=2 align=center bgcolor="#F09F46"><input type=hidden name="action" value="send"><b>请完整输入下列信息</b></td>
          </tr>
          <tr> 
            <td valign=middle bgcolor="#C68600"><b>收件人：</b></td>
            <td valign=middle bgcolor="#C68600">
            <input type=text name="touser" value="<%=request("touser")%>" size=20>
              <SELECT name=font onchange=document.messager.touser.value=options[selectedIndex].value>
              <OPTION selected value="">选择</OPTION>
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
            <td valign=top width=30% bgcolor="#C68600"><b>标题：</b></td>
            <td valign=middle bgcolor="#C68600"><input type=text name="title" size=36 maxlength=80></td>
          </tr>
          <tr> 
            <td valign=top width=30% bgcolor="#C68600"><b>内容：</b><br>Ctrl+Enter发送</td>
            <td valign=middle bgcolor="#C68600"><textarea cols=35 rows=6 name="message" title="Ctrl+Enter发送"></textarea></td>
          </tr>
          <tr> 
            <td colspan=2 align=center bgcolor="#F09F46"> 
              <input type=Submit value="发 送" name=Submit&quot;>&nbsp;<input type="reset" name="Clear" value="清 除">
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
'读取信息
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
			errmsg="<li>对不起，你是不是跑别人的信箱里了？如果不是请刷新!!!</li>"
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
          <td align=center colspan=3><b>欢迎使用短消息接收，<%=request.cookies("username")%></b></td>
        </tr>
        <tr>
          <td align=center colspan=3 bgcolor="#C68600"><a href="messager.asp?action=delete&id=<%=rs("id")%>"><img  src="images/deletepm.gif" border=0 alt="删除消息"></a> &nbsp; <a href="messager.asp?action=inbox">
             <img  src="images/inboxpm.gif" border=0 alt="收件箱"></a> &nbsp; <a href="messager.asp?action=outbox"><img  src="images/outboxpm.gif" border=0 alt="发件箱"></a>&nbsp;&nbsp; <a href="messager.asp?action=friendlist"><img border="0" src="images/friendlist.gif"></a>&nbsp; &nbsp;<a href="messager.asp?action=new"><img   src="images/newpm.gif" border=0 alt="发送消息"></a>&nbsp; &nbsp;<a href="messager.asp?action=new&touser=<%=rs("sender")%>"><img  src="images/replypm.gif" border=0 alt="回复消息"></a></td>
        </tr>
        <tr>
          <td valign=middle align=center>
<%if request("action")="outread" then%>
            在<b><%=rs("sendtime")%></b>，您发送此消息给<b><%=rs("incept")%></b>！
<%else%>
		    在<b><%=rs("sendtime")%></b>，<b><%=rs("sender")%></b>给您发送的消息！
<%end if%></font></td>
        </tr>
        <tr>
          <td valign=top align=left><b>消息标题：<%=htmlencode1(rs("title"))%></b><p><%if rs("sender")="短信精灵" then%><%=rs("content")%><%else%><%=htmlencode2(rs("content"))%><%end if%></td>
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
			response.write "<script language=javascript>window.alert('你要发给谁呀？回去写上收信人！');history.back(1);</script>"
			response.end  
	elseif request("touser")="短信精灵" then
			conn.close
		set conn=nothing
			response.write "<script language=javascript>window.alert('短信精灵是系统内置帐号，你不能发消给他，有问题请联系站长！');history.back(1);</script>"
			response.end 
	else
		incept=request("touser")
	end if
	if request("title")="" then
		conn.close
		set conn=nothing
			response.write "<script language=javascript>window.alert('标题一定要填写，回去填上标题再发送吧！');history.back(1);</script>"
			response.end
	else
		title=request("title")
	end if
	if request("message")="" then
		conn.close
		set conn=nothing
			response.write "<script language=javascript>window.alert('短信息的内容不能为空呀，写上两句话吧！');history.back(1);</script>"
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
			response.write "<script language=javascript>window.alert('收信人不是本站用户！');history.back(1);</script>"
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
			msg=msg+"<br>"+"<li><b>恭喜您，发送短信息成功。</b><br>发送的消息同时保存在您的发件箱。"
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
			response.write "<script language=javascript>window.alert('短信息删除失败！');history.back(1);</script>"
			response.end 
	else
		msg=msg+"<br>"+"<li>短信息成功删除！"
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
			response.write "<script language=javascript>window.alert('短信息删除失败！');history.back(1);</script>"
			response.end 
	else
		msg=msg+"<br>"+"<li>短信息全部成功删除！"
		call success()
	end if
end sub

sub friendadd
touser=checkStr(request("touser"))
REM 判断用户是否存在
sql="select username from [user] where username='"&touser&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
			response.write "<script language=javascript>window.alert('很抱歉，本站没有这个用户！');history.back(1);</script>"
			response.end 
	exit sub
end if
if touser=request.cookies("username") then
			response.write "<script language=javascript>window.alert('不能添加你自己为好友！');history.back(1);</script>"
			response.end 
	exit sub
end if
if request("touser")="短信精灵" then
			conn.close
		set conn=nothing
			response.write "<script language=javascript>window.alert('短信精灵是系统内置帐号，你不能加他为好友，有问题请联系站长！');history.back(1);</script>"
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
'删除好友
sub frienddel
delid=request("id")
if delid="" or isnull(delid) then
			response.write "<script language=javascript>window.alert('你的操作错误或程序出现严重错误！');history.back(1);</script>"
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
              <td width="100%" bgcolor="#" >成功：短信息</td>
            </tr>
            <tr> 
              <td width="100%"><%=msg%>　</td>
            </tr>
            <tr align="center"> 
              <td width="100%"><a href="javascript:history.go(-1)"> << 返回上一页</a> </td>
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
'新信息提示
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
              <td width="100%"  bgcolor="#F09F46">短消息通知</td>
            </tr>
            <tr> 
              <td width="100%" align=center bgcolor="#C68600"><br>
                <a href=messager.asp?action=inbox><img src="images/newmail.gif" border=0>有新的短消息</a><br>
                <br>
                <a href=messager.asp?action=inbox>按此查看</a><br><br>
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
'好友列表
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
              <td align=center width="398" bgcolor="#F09F46"><b>用户<%=request.cookies("username")%>的好友列表</b></td>
            </tr>
            <tr>
              <td align=center bgcolor="#C68600" width="398"><a href="messager.asp?action=inbox"><img  src="images/inboxpm.gif" border=0 alt="收件箱"></a> &nbsp; <a href="messager.asp?action=outbox">
                 <img  src="images/outboxpm.gif" border=0 alt="发件箱"></a>&nbsp;&nbsp; <a href="messager.asp?action=friendlist"><img border="0" src="images/friendlist.gif"></a> &nbsp; <a href="messager.asp?action=new"><img  src="images/newpm.gif" border=0 alt="发送消息"></a></td>
            </tr>
            <tr>
              <td align=center bgcolor="#C68600" width="100%">

<div align="center">
  <center>
<table cellpadding=0 cellspacing=0 width="398" border="1" bordercolor="#000000" style="border-collapse: collapse">
            <tr>
                <td width="114" align="center">姓名</td> 
                <td width="121" align="center">邮件</td> 
                <td width="92" align="center">OICQ</td>
                <td width="33" align="center">短信</td> 
                <td width="32" align="center">删除</td></tr>
<%
	sql="select F.*,U.id,U.email,U.homepage,U.oicq from Friend F inner join [user] U on F.F_Friend=U.username where F.F_username='"&request.cookies("username")&"' order by F.f_addtime desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
                <tr>
                <td class=tablebody1 align=center valign=middle colspan=5 width="396">您的好友列表中没有任何内容。</td>
                </tr>
<%else%>
<%do while not rs.eof%>
                <tr>
                    <td align=center valign=middle class=tablebody1 width="114"><a href=### onclick="javascript:open_window('UserInfo.asp?id=<%=rs("id")%>','UserInfo','scrollbars=no,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,copyhistory=no,top=10,left=10,width=435,height=448')"><%=htmlencode(rs("F_friend"))%></a>　</td>
                    <td align=center valign=middle class=tablebody1 width="121"><a href="mailto:<%=htmlencode(rs("email"))%>"><%=htmlencode(rs("email"))%></a>　</td>
                    <td align=center class=tablebody1 width="92"><%=htmlencode(rs("oicq"))%>　</td>
                    <td align=center class=tablebody1 width="33"><a href="messager.asp?action=new&touser=<%=htmlencode(rs("f_friend"))%>">发送</a></td>
                <td align=center class=tablebody1 width="32">
                <a href="messager.asp?action=frienddel&id=<%=rs("f_id")%>">删除</a></td>
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
              <input type=text name="touser" value="输入想要添加为好友的用户名" size=28>
              <input type=submit value="添  加" name=Submit></td>
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