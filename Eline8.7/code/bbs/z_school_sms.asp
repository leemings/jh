<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!-- #include file="inc/z_school_char.asp" -->
<%
'=========================================================
' File: z_school_sms.asp
' Version:5.0
' Date: 2003-1-13
' Script Written by wxzlxl http://www.zmcn.com QQ:628122
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
%>
<html>
<head>
<title>全班短消息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>
<body <%=Forum_body(11)%>" onkeydown="if(event.keyCode==13 && event.ctrlKey)messager.submit()">
<%
dim msg
dim abgcolor
dim username
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	founderr=true
elseif BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
Errmsg=Errmsg+"<br><li> 错误的同学录参数！请确认您是从有效的连接进入。"
	founderr=true
elseif cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本同学录为认证同学录，请<a href=login.asp>登录</a>并确认您是本班同学。"
		founderr=true
	else
		if chkschoollogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>你不是本班同学，请先<a href=z_school_inclass.asp?boardid="&boardid&">加入本班</a>。"
		founderr=true
		end if
	end if
end if
stats="全班短消息"
set rs=conn.execute("select boarduser from board where boardid="&boardid)
dim boarduser,sss
sss=""
if rs(0)<>"" then
boarduser=split(replace(rs(0)," ",""),",")
for i = 0 to ubound(boarduser)
sss=sss&"<OPTION value='"&boarduser(i)&"'>"&boarduser(i)&"</OPTION>"
next
else
Errmsg=Errmsg+"<br>"+"<li>该同学录还没有同学！"
founderr=true
end if
boarduser=rs(0)
set rs=nothing
if founderr then
	call dvbbs_error()
else
	select case request("action")
	case "send"
		call savemsg()
	case else
call sendmsg()
	end select
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
end if
%>
<form action="z_school_sms.asp?BoardID=<%=BoardID%>" method=post name=messager>
<input type=hidden name="action" value="send">
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
          <tr> 
            <th colspan=3>发送短消息（请输入完整信息）</td>
          </tr>
          <tr> 
            <td class=tablebody1 valign=middle><b>收件人：</b></td>
            <td class=tablebody1 valign=middle>
              <input type=text name="touser" value="<%=request("touser")%>" size=50>
              <SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)>
              <OPTION selected value="">全班</OPTION><%=sss%>"></SELECT>
            </td>
          </tr>
          <tr> 
            <td class=tablebody1 valign=top width=15%><b>标题：</b></td>
            <td class=tablebody1 valign=middle>
              <input type=text name="title" size=70 maxlength=80 value="<%=title%>">
            </td>
          </tr>
          <tr> 
            <td class=tablebody1 valign=top width=15%><b>内容：</b></td>
            <td  class=tablebody1 valign=middle>
              <textarea cols=70 rows=6 name="message" title="Ctrl+Enter发送"><%if request("id")<>"" then%>
====== 在 <%=sendtime%> 您来信中写道： ======
<%=server.htmlencode(content)%>
=====================================
<%end if%></textarea>
            </td>
          </tr>
          <tr> 
            <td  class=tablebody1 colspan=2>
<b>说明</b>：<br>
① 您可以使用<b>Ctrl+Enter</b>键快捷发送短信<br>
② 可以用英文状态下的逗号将用户名隔开实现群发，最多<b><%=GroupSetting(33)%></b>个用户<br>
③ 标题最多<b>50</b>个字符，内容最多<b><%=GroupSetting(34)%></b>个字符<br>
            </td>
          </tr>
          <tr> 
            <td  class=tablebody2 valign=middle colspan=2 align=center> 
              <input type=Submit value="发送" name=Submit>
              &nbsp; 
              <input type=Submit value="保存" name=Submit>
              &nbsp; 
              <input type="reset" name="Clear" value="清除">
              &nbsp; 
<%if request("reaction")="chatlog" then%>
              <input type=button value="关闭聊天记录" name="chatlog" onclick="location.href='?action=new&id=<%=request("id")%>&touser=<%=request("touser")%>'">
<%else%>
              <input type=button value="查看聊天记录" name="chatlog" onclick="location.href='?action=new&id=<%=request("id")%>&touser=<%=request("touser")%>&reaction=chatlog'">
<%end if%>
              &nbsp; 
              <input type="button" name="close" value="关闭" onclick="window.close()">
            </td>
          </tr>
<%if request("reaction")="chatlog" then%>
          <tr> 
            <th colspan=3>我与<%=request("touser")%>的聊天记录</th>
          </tr>
<%if membername=request("touser") then%>
          <tr> 
            <td class=tablebody1 colspan=3>自己跟自己的聊天记录没什么好看的：）</td>
          </tr>
<%else%>
<%
	set rs=server.createobject("adodb.recordset")
	sql="select * from message where ((sender='"&trim(membername)&"' and incept='"&replace(request("touser"),"'","")&"') or (sender='"&replace(request("touser"),"'","")&"' and incept='"&membername&"')) and delS=0 order by sendtime desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
          <tr> 
            <td class=tablebody1 colspan=3>还没有任何聊天记录！</td>
          </tr>
<%
	else
	do while not rs.eof
%>
                <tr>
                    <td class=tablebody2 height=25 colspan=3>
<%if rs("sender")=membername then%>
                    在<b><%=rs("sendtime")%></b>，您发送此消息给<b><%=htmlencode(rs("incept"))%></b>！
<%else%>
		    在<b><%=rs("sendtime")%></b>，<b><%=htmlencode(rs("sender"))%></b>给您发送的消息！
<%end if%></td>
                </tr>
                <tr>
                    <td  class=tablebody1 valign=top align=left colspan=3>
                    <b>消息标题：<%=htmlencode(rs("title"))%></b><hr size=1>
                    <%=dvbcode(rs("content"),4,3)%>
		    </td>
                </tr>
<%
	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
%>
<%end if%>
<%end if%>
        </table>
</form>
<%
end sub


'1111111
sub savemsg()
stats="发送短信成功"
dim incept,title,message,subtype
if request("touser")="" then
incept=boarduser
else
incept=request("touser")
end if
incept=CheckStr(incept)
incept=split(incept,",")
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
for i=0 to ubound(incept)
sql="select username from [user] where username='"&replace(incept(i),"'","")&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>论坛没有这个用户，看看你的发送对象写对了嘛？"
	founderr=true
	exit sub
end if
set rs=nothing
if request("Submit")="发送" then
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,1)"
	subtype="已发送信息"
elseif request("Submit")="保存" then
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,0)"
	subtype="发件箱"
else
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,1)"
	subtype="已发送信息"
end if
conn.execute(sql)
next
sucmsg=sucmsg+"<br>"+"<li><b>恭喜您，发送短信息成功。</b><br>发送的消息同时保存在您的"&subtype&"中。"
call dvbbs_suc()
end sub
%>
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