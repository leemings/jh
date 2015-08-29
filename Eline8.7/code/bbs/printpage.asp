<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
'=========================================================
' File: printpage.asp
' Version:5.0
' Date: 2002-9-10
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim urs,usql
dim rsboard,boardsql
dim announceid
dim username
dim rootid
dim topic
dim abgcolor
dim postbuyuser
dim replyid
abgcolor="#FFFFFF"
Dim TotalUseTable
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
else
	AnnounceID=request("id")
end if

if cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
	else
		if chkboardlogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
		end if
	end if
end if

if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>您没有权限进入隐含论坛！"
		founderr=true
	end if
end if

'###################特殊版面修改开始(asilas制作)##################
if cint(Board_Setting(51))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为特殊论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
	else
		if chkviplogin(membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>本论坛版面为<font color=red>VIP会员专用版面</font>，请确认您的属性是否符合。"
		founderr=true
		end if
	end if
end if
if cint(Board_Setting(52))<>0 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为特殊论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
	else
		dim sexshow
		if cint(Board_Setting(52))=1 then
		sexshow="女生"
		elseif cint(Board_Setting(52))=2 then
		sexshow="男生"
		end if
		if chksexlogin(cint(Board_Setting(52)),membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>本论坛版面为<font color=red>"&sexshow&"论坛版面</font>，请确认您的性别是否符合。"
		founderr=true
		end if
	end if
end if
'####################特殊版面修改结束(asilas制作)#################

if foundErr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call announceinfo()
	if founderr then call dvbbs_error()
end if
call footer()

sub announceinfo()
    set rs=conn.execute("select title,PostTable,LockTopic,PostUserName from topic where topicID="&AnnounceID)
	if not(rs.bof and rs.eof) then
		if rs("locktopic")=2 then
			ErrMsg=ErrMsg+"<br>"+"<li>该帖子已经被管理员删除！</li>"
			founderr=true
			exit sub
		end if
		TotalUseTable=rs(1)
		topic=rs(0)
		if lcase(trim(rs("PostUserName")))<>lcase(trim(membername)) then
			if Cint(GroupSetting(2))=0 then
				Errmsg=Errmsg+"<br>"+"<li>您没有浏览在本论坛查看其他人发布的帖子的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
				founderr=true
				exit sub
			end if
		end if
	else
		foundErr = true
		ErrMsg=ErrMsg+"<br>"+"<li>您指定的帖子不存在</li>"
		exit sub
	end if
	rs.close
%>
<HTML><HEAD><TITLE><%=Forum_info(0)%>--显示帖子</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<!--#include file="inc/Forum_css.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY bgcolor="#ffffff" alink="#333333" vlink="#333333" link="#333333" topmargin="10" leftmargin="10">
<TABLE border=0 width="<%=Forum_body(12)%>" align=center>
  <TBODY>
  <TR>
    <TD valign=middle align=top>
<b>以文本方式查看主题</b><br><br>
-&nbsp;&nbsp;<b><%=Forum_info(0)%></b>&nbsp;&nbsp;(<%=Forum_info(1)%>index.asp)<br>
--&nbsp;&nbsp;<b><%=boardtype%></b>&nbsp;&nbsp;(<%=Forum_info(1)%>list.asp?boardid=<%=boardid%>)<br>
----&nbsp;&nbsp;<b><%=htmlencode(topic)%></b>&nbsp;&nbsp;(<%=Forum_info(1)%>dispbbs.asp?boardid=<%=boardid%>&rootid=<%=rootid%>&id=<%=announceid%>)
      </TD></TR></TBODY></TABLE> 
<br>
<hr>
<%
	Rs.open "Select b.UserName,b.Topic,b.dateandtime,b.body,u.UserGroupID,b.postbuyuser from "&TotalUseTable&" b inner join [user] u on b.PostUserID=u.userid where b.boardid="&boardid&" and b.rootid="&Announceid&" order by b.announceid",conn,1,1
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>没有找到相关帖子。"
		exit sub
	end if
	do while not rs.eof
	postbuyuser=rs("postbuyuser")
	username=rs(0)
%>
<TABLE border=0 width="750" align=center>
  <TBODY>
  <TR>
    <TD valign=middle align=top>
--&nbsp;&nbsp;作者：<%=rs("username")%><br>
--&nbsp;&nbsp;发布时间：<%=rs("dateandtime")%><br><br>
--&nbsp;&nbsp;<%=htmlencode(rs("topic"))%><br>
<%=dvbcode(rs("body"),rs("UserGroupID"),1)%>
	<hr>
    </TD></TR></TBODY></TABLE> 
<%
          rs.movenext
        loop	
	rs.close
	set rs=nothing
	call activeonline()
end sub    
%>