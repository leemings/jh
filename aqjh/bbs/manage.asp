<!-- #include file="setup.asp" -->
<!-- #include file="jhconst.asp" -->
<%
if Request.Cookies("username")=empty then error2("请登录后再进行操作！")

if Request.ServerVariables("request_method") <> "POST" then
response.write "<form name=BBSxpPOST method=post action=manage.asp?"&Request.ServerVariables("Query_String")&"></form><SCRIPT>if(confirm('您确定要执行该操作?')){returnValue=BBSxpPOST.submit()}else{returnValue=history.back()}</SCRIPT>"
htmlend
end if


top


id=int(Request("id"))
sql="Select * From [forum] where ID="&ID&""
rs.Open sql,Conn,1
forumid=rs("forumid")
ReList=rs("ReList")
rs.close

if membercode > 3 then
pass=1
elseif instr("|"&Conn.Execute("Select moderated From [bbsconfig] where id="&forumid&" ")(0)&"|","|"&Request.Cookies("username")&"|")>0 then
pass=1
end if

if pass<>1 then error("<li>您的权限不够")


username=Conn.Execute("Select username From [forum] where id="&id&"")(0)

select case Request("menu")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "top"
if membercode > 3 then
sql="Select * From [forum] where id="&id
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 topic=rs("topic")
else
 forum_username="未知人"
 topic="未知"
end if
rs.close
'奖励
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update 用户 set 泡豆点数=泡豆点数+"&bbs_add4&" where 姓名='"&forum_username&"'")
connjh.close
set connjh=nothing
'奖励
mess="<font color=000099><b>〖论坛消息〗</b></font>论坛管理<font color=#cc0000>"&jhname&"</font><font color=006600>设置了["&forum_username&"]的主题["&topic&"]总固顶，特奖励豆点"&bbs_add4&"个</font>"
call showchat (mess)
conn.execute("update [forum] set toptopic=2 where id="&id&"")
succtitle="总置顶主题成功"
else
error("<li>您的权限不够")
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "untop"
if membercode > 3 then
conn.execute("update [forum] set toptopic=0 where id="&id&"")
succtitle="取消总置顶主题成功"
else
error("<li>您的权限不够")
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "movenew"
sql="Select * From [forum] where id="&id
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 topic=rs("topic")
else
 forum_username="未知人"
 topic="未知"
end if
rs.close
conn.execute("update [forum] set lasttime='"&now()&"' where id="&id&"")
'奖励
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update 用户 set 泡豆点数=泡豆点数+"&bbs_add3&" where 姓名='"&forum_username&"'")
connjh.close
set connjh=nothing
'奖励
mess="<font color=000099><b>〖论坛消息〗</b></font>论坛管理<font color=#cc0000>"&jhname&"</font><font color=006600>拉前了["&forum_username&"]的主题["&topic&"]，并奖励豆点"&bbs_add3&"个</font>"
call showchat (mess)
succtitle="拉前主题成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "move"
if Request("moveid")="" then error("<li>您没有选择要将主题移动哪个论坛")
if Conn.Execute("Select pass From [bbsconfig] where id="&Request("moveid")&"")(0)=4 then error("<li>目标论坛为授权发帖状态")
sql="Select * From [forum] where id="&id
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 topic=rs("topic")
else
 forum_username="未知人"
 topic="未知"
end if
rs.close
'惩罚
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update 用户 set 泡豆点数=泡豆点数-"&bbs_del4&" where 姓名='"&forum_username&"'")
connjh.close
set connjh=nothing
'惩罚
mess="<font color=000099><b>〖论坛消息〗</b></font>论坛管理<font color=#cc0000>"&jhname&"</font><font color=006600>移动了["&forum_username&"]的主题["&topic&"]，由于发贴不当,因此扣除豆点"&bbs_del4&"个</font>"
call showchat (mess)
conn.execute("update [forum] set forumid="&Request("moveid")&",toptopic=0,goodtopic=0,locktopic=0 where id="&id&"")
succtitle="移动主题成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "deltopic"
if isnumeric(""&Request("retopicid")&"") then
sql="Select * From [reforum] where topicid="&id&" and id="&Request("retopicid")
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 forum_topicid=rs("topicid")
else
 forum_username="未知人"
 forum_topicid="未知"
end if
rs.close
conn.execute("delete from [reforum"&ReList&"] where topicid="&id&" and id="&Request("retopicid")&"")
conn.execute("update [forum] set replies=replies-1 where id="&id&"")
succtitle="删除回贴成功"
mess="<font color=000099><b>〖论坛消息〗</b></font>论坛管理<font color=#cc0000>"&jhname&"</font><font color=006600>删除["&forum_username&"]在主题ID["&forum_topicid&"]的贴子，并扣取豆点"&bbs_del2&"个</font>"
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update 用户 set 泡豆点数=泡豆点数-"&bbs_del2&" where 姓名='"&forum_username&"'")
connjh.close
set connjh=nothing
call showchat (mess)
else
sql="Select * From [forum] where id="&id&" and deltopic=0"
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_topic=rs("topic")
 forum_username=rs("username")
else
 forum_username="未知人"
 forum_topic="未知"
end if
rs.close
conn.execute("update [user] set deltopic=deltopic+1 where username='"&username&"'")
conn.execute("update [forum] set toptopic=0,deltopic=1,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and deltopic=0")
succtitle="删除主题成功"
mess="<font color=000099><b>〖论坛消息〗</b></font>论坛管理<font color=#cc0000>"&session("aqjh_name")&"</font><font color=006600>删除["&forum_username&"]的主题["&forum_topic&"]，并扣取豆点"&bbs_del1&"个</font>"
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update 用户 set 泡豆点数=泡豆点数-"&bbs_del1&" where 姓名='"&forum_username&"'")
connjh.close
set connjh=nothing
call showchat (mess)
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "goodtopic"
if Conn.Execute("Select goodtopic From [forum] where id="&id&" ")(0)=1 then error("<li>此帖子已经加入精华区了，无需重复添加")
sql="Select * From [forum] where id="&id
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 topic=rs("topic")
else
 forum_username="未知人"
 topic="未知"
end if
rs.close
'奖励
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update 用户 set 泡豆点数=泡豆点数+"&bbs_add6&" where 姓名='"&forum_username&"'")
connjh.close
set connjh=nothing
'奖励
mess="<font color=000099><b>〖论坛消息〗</b></font>论坛管理<font color=#cc0000>"&jhname&"</font><font color=006600>设置了["&forum_username&"]的主题["&topic&"]为精华贴，特奖励豆点"&bbs_add6&"个</font>"
call showchat (mess)
conn.execute("update [forum] set goodtopic=1 where id="&id&"")
conn.execute("update [user] set goodtopic=goodtopic+1 where username='"&username&"'")
succtitle="加入精华区成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "delgoodtopic"
if Conn.Execute("Select goodtopic From [forum] where id="&id&" ")(0)=0 then error("<li>此帖子已经移出精华区了")
conn.execute("update [forum] set goodtopic=0 where id="&id&"")
conn.execute("update [user] set goodtopic=goodtopic-1 where username='"&username&"'")
succtitle="移出精华区成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "toptopic"
sql="Select * From [forum] where id="&id
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 topic=rs("topic")
else
 forum_username="未知人"
 topic="未知"
end if
rs.close
'奖励
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update 用户 set 泡豆点数=泡豆点数+"&bbs_add5&" where 姓名='"&forum_username&"'")
connjh.close
set connjh=nothing
'奖励
mess="<font color=000099><b>〖论坛消息〗</b></font>论坛管理<font color=#cc0000>"&jhname&"</font><font color=006600>设置了["&forum_username&"]的主题["&topic&"]版块固顶，特奖励豆点"&bbs_add5&"个</font>"
call showchat (mess)
conn.execute("update [forum] set toptopic=1 where id="&id&"")
succtitle="置顶主题成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "deltoptopic"
conn.execute("update [forum] set toptopic=0 where id="&id&"")
succtitle="取消置顶主题成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "locktopic"
sql="Select * From [forum] where id="&id
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 topic=rs("topic")
else
 forum_username="未知人"
 topic="未知"
end if
rs.close
'惩罚
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update 用户 set 泡豆点数=泡豆点数-"&bbs_del3&" where 姓名='"&forum_username&"'")
connjh.close
set connjh=nothing
'惩罚
mess="<font color=000099><b>〖论坛消息〗</b></font>论坛管理<font color=#cc0000>"&jhname&"</font><font color=006600>关闭了["&forum_username&"]的主题["&topic&"]，并扣除豆点"&bbs_del3&"个</font>"
call showchat (mess)
conn.execute("update [forum] set locktopic=1 where id="&id&"")
succtitle="关闭主题成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "dellocktopic"
conn.execute("update [forum] set locktopic=0 where id="&id&"")
succtitle="开放主题成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "lookip"
if isnumeric(""&Request("retopicid")&"") then
sql="Select * From [reforum"&ReList&"] where id="&Request("retopicid")&""
else
sql="Select * From [forum] where id="&id&""
end if

rs.Open sql,Conn,1
username=rs("username")
posttime=rs("posttime")
postip=rs("postip")
rs.close
%>


<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 查看IP</td>
</tr>
</table>
<br>
<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>
<tr>
<td width="328" height="25" align="center" class=a1 colspan="2">
查看IP
</td></tr><tr>
<td height="7" width="164" valign="top" align="center" class=a3>
用户名</td>
<td height="7" width="164" valign="top" align="center" class=a3>
<%=username%></td></tr><tr>
<td height="6" width="164" valign="top" align="center" class=a3>
时间</td>
<td height="6" width="164" valign="top" align="center" class=a3>
<%=posttime%></td></tr><tr>
<td height="6" width="164" valign="top" align="center" class=a3>
IP地址</td>
<td height="6" width="164" valign="top" align="center" class=a3>
<%=postip%></td></tr></table>

<br>
<center>
<a href=ShowPost.asp?id=<%=id%>>BACK</a><br>
<%
htmlend

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
end select
'聊天室显示
Sub showchat(mess)
says="<bgsound src=wav/xintie.wav loop=1><img src=img/xintie.gif>"&mess
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & session("aqjh_name") & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
addmsg saystr
end sub
Function Yushu1(a)
 Yushu1=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu1(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
'聊天室显示
if succtitle="" then error("<li>无效命令")

log(""&succtitle&"，主题ID："&id&"")

message="<li>"&succtitle&"<li><a href=ShowForum.asp?forumid="&forumid&">返回论坛</a><li><a href=Default.asp>返回社区首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?forumid="&forumid&">")
%>