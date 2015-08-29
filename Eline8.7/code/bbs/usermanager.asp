<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="z_visual_const.asp" -->
<%Response.Buffer = true%>
<%
'=========================================================
' File: usermanager.asp
' Version:5.0
' Date: 2002-9-11
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

stats="控制面板首页"
call nav()
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call head_var(0,0,membername & "的控制面板","usermanager.asp")
	call main()
	call activeonline()
end if
call footer()
sub main
	dim userclass
	set rs=server.createobject("adodb.recordset")
	sql="select * from [User] where userid="&userid
	rs.open sql,conn,1,1
	userclass=rs("userclass")
%>
<!--#include file="z_controlpanel.asp"-->
<br>
<%
response.write "<table cellpadding=0 cellspacing=6 width="&Forum_body(12)&" align=center >"&_
               "<tr  align=center ><td  width=28% valign=top>"&_
               "<table align=center style=""width:100%"" height=100% cellspacing=1 cellpadding=6 class=tableborder1>"&_
               "<tr><th height=25>用户形象</th></tr>"&_
               "<tr align=center><td class=tablebody1>"
call ShowMyselfVisual(140,"tablebody1")
response.write "</td></tr>"&_
               "<tr><th height=25>基本信息</th></tr>"&_
               "<tr><td align=left class=tablebody1 valign=top>" 
if trim(rs("title"))<>"" then
response.write "用户头衔： "&htmlencode(RS("title"))&"<br>"
end if

response.write "用户等级： "& rs("userclass") &"<br>"

if rs("UserGroup")<>"" then
response.write "用户门派： "&rs("UserGroup")&"<br>"
end if

response.write "用户财富： "&rs("userwealth")&"<br>"&_
               "用户经验： "&rs("userEP")&"<br>"&_
               "用户魅力： "&rs("userCP")&"<br>"&_
               "精华帖数： "&rs("userisbest")&"<br>"&_
               "帖数总数： "&rs("article")&"<br>"&_
               "注册时间： "&rs("addDAte")&"<br>"&_
               "登录次数： "&rs("logins")&"<br>"
        rs.close
        set rs=nothing

response.write"</td></tr></table><br>"
call friend()

response.write "</td>"&_
              "<td valign=top>"&_
              "<table cellpadding=3 cellspacing=1 style=""width:100%"" height=29 align=center  class=tableborder1>"&_
              "<tr>"&_
              "<th height=25 align=left>-=> 论坛短信息"&_
              "</td></tr><tr><td class=tablebody1>"
if newincept() >0 then
        response.write"目前您有<font color="""&Forum_body(8)&"""><b> ["&newincept()&"] </b></font>条的短消息。"
        else
        response.write"目前您没有新的短消息,"
        end if
response.write"<a href=usersms.asp?action=inbox><font color="""&Forum_body(8)&""">收件箱</font></a>中共有 <b>["&cint(incept())&"]</b> 条信息，<a href='usersms.asp?action=issend'><font color="""&Forum_body(8)&""">发件箱</font></a>中共有 <b>["&allsend()&"]</b> 条信息对方未查阅。<br>"&_
              "</td></tr></table><br>"&_
              "<table cellpadding=3 cellspacing=1 style=""width:100%"" align=center class=tableborder1>"&_
              "<tr>"&_
              "<th colspan=5 height=25 align=left>-=> 最新收到的短讯</th></tr>"&_
            "<tr>"&_
                "<td align=center valign=middle width=30 class=tabletitle2><b>状态</b></td>"&_
                "<td align=center valign=middle width=100 class=tabletitle2><b>发件人</b></td>"&_
                "<td align=center valign=middle width=*  class=tabletitle2><b>主题</b></td>"&_
                "<td align=center valign=middle width=120 class=tabletitle2><b>日期</b></td>"&_
                "<td align=center valign=middle width=60 class=tabletitle2><b>大小</b></td>"&_
            "</tr>"

	set rs=server.createobject("adodb.recordset")
	sql="select top 5 * from message where incept='"&trim(membername)&"' and issend=1 and delR=0 order by flag,sendtime desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then

response.write"<tr><td class=tablebody1 align=center valign=middle colspan=6>您的收件箱中没有任何内容。</td></tr>"
        else
		dim tablebody
do while not rs.eof
response.write "<tr>"
if rs("flag")=0 then
tablebody="tablebody2"
else
tablebody="tablebody1"
end if
response.write"<td align=center valign=middle class="&tablebody&">"

if rs("flag")=0 then
response.write"<img src=pic/m_news.gif>"
else
response.write"<img src="&Forum_info(7)&"m_olds.gif>"
end if
response.write"</td>"&_
              "<td align=center valign=middle class="&tablebody&">"
if rs("flag")=0 then
response.write"<b>"
end if
response.write"<a href=dispuser.asp?name="&htmlencode(rs("sender"))&" target=_blank>"&htmlencode(rs("sender"))&"</a></td>"&_
              "<td align=left class="&tablebody&"><a href=usersms.asp?action=read&id="&rs("id")&"&sender="&rs("sender")&" >"
if rs("flag")=0 then
response.write"<b>"
end if
response.write""&htmlencode(rs("title"))&"</a></td>"&_
               "<td class="&tablebody&">"
if rs("flag")=0 then
response.write"<b>"
end if
response.write""&rs("sendtime")&"</font></td>"&_
               "<td class="&tablebody&">"
if rs("flag")=0 then
response.write"<b>"
end if
response.write""&len(rs("content"))&"Byte</td></tr>"

	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
    response.write"</td></tr></table><br>"


	dim F_Type,F_Typename
'---------------------
response.write"<table cellpadding=3 cellspacing=1 style=""width:100%"" align=center class=tableborder1>"&_
						"<tr>"&_
						"<th colspan=5 height=25 align=left>-=> 最新上传文件</th></tr>"&_
						"<tr>"&_
						 "<td align=center valign=middle width=30 class=tabletitle2><b>属性</b></td>"&_
						 "<td align=center valign=middle width=100 class=tabletitle2><b>大小</b></td>"&_
						 "<td align=center valign=middle width=* class=tabletitle2><b>文件</b></td>"&_
						 "<td align=center valign=middle width=120 class=tabletitle2><b>日期</b></td>"&_
						"<td align=center valign=middle width=60 class=tabletitle2><b>类型</b></td>"&_
						"</tr>"
'---------------------
set rs=conn.execute("select top 5  F_ID,F_Filename,F_FileType,F_Type,F_Flag,F_FileSize,F_AddTime from [DV_Upfile] where   F_UserID="&userid&" order by F_ID desc ")
if rs.eof and rs.bof then
		response.write"<tr><td class=tablebody1 align=center valign=middle colspan=5>您的文件库中没有任何内容。</td></tr>"
else
		do while not rs.eof
		F_Type=rs("F_Type")
		if F_Type=1 then
			F_Typename="图片集"
		elseif F_Type=2 then
			F_Typename="FLASH集"
		elseif F_Type=3 then
			F_Typename="音乐集"
		elseif F_Type=4 then
			F_Typename="电影集"
		else
			F_Typename="文件集"
		end if
'---------------------
response.write		"<tr>"&_
							"<td align=center valign=middle  class=tablebody1><img src='images/files/"&rs("F_FileType")&".gif' border=0></td>"&_
							"<td align=left valign=middle  class=tablebody1>"&rs("F_FileSize")&" Byte</td>"&_
							 "<td align=left valign=middle  class=tablebody1><a href=""myfile.asp"" >"&rs("F_Filename")&"</a></td>"&_
							"<td align=left valign=middle  class=tablebody1>"&rs("F_AddTime")&"</td>"&_
							"<td align=center valign=middle  class=tablebody1><b>"&F_Typename&"</b></td>"&_
							"</tr>"
'---------------------
rs.movenext
loop
end if
rs.close
set rs=nothing
response.write		"</table><br>"&_
							"<table cellpadding=3 cellspacing=1 style=""width:100%"" align=center class=tableborder1>"&_
							"<tr>"&_
							"<th height=25 align=left>-=> 最近发表的文章</th></tr>"

dim topic,srs
set srs=conn.execute("select top 5 rootid,boardid,dateandtime,topic,expression,body,announceid from "&NowUseBBS&" where PostUserID="&userid&" and not locktopic=2 order by announceid desc")
do while not srs.eof
topic=replace(srs(3)," ","")
          if topic<>"" then
          topic=topic
          else
          topic=srs(5)
          topic=replace(topic,chr(13),"")
          topic=replace(topic,chr(10),"")
          end if
          if len(topic)>30 then
          topic=left(topic,30)&"..."
          end if
          topic=server.htmlencode(topic)
response.write" <tr>"&_ 
						"<td align=left class=tablebody1>&nbsp;□　&nbsp;<a href=dispbbs.asp?boardid="&srs(1)&"&id="&srs(0)&"&replyid="&srs(6)&"#"&srs(6)&">"&htmlencode(topic)&"</a>&nbsp;--&nbsp;"&srs(2)&"</td>"&_
						"</tr>"&_

srs.movenext
loop
srs.close
set srs=nothing

response.write"</table><br></td></tr></table>"
end sub

sub friend()
dim frs,srs,OnlineTime
response.write"<table align=center style=""width:100%"" height=100% cellspacing=1 cellpadding=6 class=tableborder1>"&_
						"<tr><th height=25>好友在线</th></tr>"&_
						"<tr align=center><td class=tablebody1 align=left>"

set frs=conn.execute("select F_friend from Friend  where F_username='"&trim(membername)&"' and F_type=0 order by F_id desc")
if frs.eof and frs.bof then
		response.write"快添加您的好友吧！"
else
		do while not frs.eof
	if boardmaster or master then
	set Srs=conn.execute("select startime from online where username='"&frs(0)&"' ")
	else
	set Srs=conn.execute("select startime from online where username='"&frs(0)&"' and userhidden=2")
	end if
	if Srs.eof and Srs.bof then
	OnlineTime="[离线]"
	else
	OnlineTime="[<font color=red>在线：" & datediff("n",Srs(0),Now()) & "Mins</font>]"
	end if
	Srs.close
	set Srs=nothing
	response.write"<a href=""javascript:openScript('messanger.asp?action=new&touser="&htmlencode(frs(0))&"',500,400)"" ><img src="""&Forum_info(7)&Forum_boardpic(9)&""" alt='给好友发送短讯' border=0></a>"
		response.write"&nbsp;<a href=dispuser.asp?name="&htmlencode(frs(0))&" >"
		response.write	  frs(0)&"</a> "&OnlineTime&" <br>"

		frs.movenext
		loop
end if
frs.close
set frs=nothing
response.write"</td></tr>"&_
						"<tr><td height=25 class=tablebody2>＊点击图标给好友发送短讯！</td></tr>"&_
						"</table>"
end sub

REM 统计已发出新短信
function allsend()
        rs=conn.execute("Select Count(id) From Message Where flag=0 and issend=1 And sender='"& membername &"'")
        allsend=rs(0)
	set rs=nothing
	if isnull(allsend) then allsend=0
end function
REM 统计收件箱中的短信
function incept()
        incept=0
        rs=conn.execute("Select Count(id) From Message Where  issend=1 and delR=0 And incept='"& membername &"'")
        incept=rs(0)
	set rs=nothing
	if isnull(incept) then incept=0
end function
%>