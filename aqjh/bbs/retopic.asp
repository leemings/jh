<!-- #include file="setup.asp" -->
<!-- #include file="jhconst.asp" -->
<%
top

if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

DetectPost

id=int(Request("id"))

sql="Select * From [forum] where ID="&ID&""
rs.Open sql,Conn,1
if rs("locktopic")=1 then error("<li>此主题已经关闭，不接受新的回复")
forumid=rs("forumid")
ReList=rs("ReList")
rs.close

username=Request.Cookies("username")
topic=HTMLEncode(Request("topic"))
content=ContentEncode(Request.Form("content"))



if badwords<>empty then
filtrate=split(badwords,"|")
for i = 0 to ubound(filtrate)
topic=ReplaceText(topic,""&filtrate(i)&"",string(len(filtrate(i)),"*"))
content=ReplaceText(content,""&filtrate(i)&"",string(len(filtrate(i)),"*"))
next
end if

If not conn.Execute("Select username From [prison] where username='"&username&"'" ).eof Then error("<li>您被关进<a href=prison.asp>监狱</a>")


sql="select * from bbsconfig where id="&forumid&""
Set Rs=Conn.Execute(sql)
bbsname=rs("bbsname")
logo=rs("logo")
followid=rs("followid")
rs.close

if membercode>1 or instr("|"&moderated&"|","|"&Request.Cookies("username")&"|")>0 then UserPopedomPass=1


if Request.ServerVariables("request_method") = "POST" and Request("preview")="" then

if Len(content)<2 then message=message&"<li>文章内容不能小于 2 字符"

if Request.Cookies("posttime")<>empty and UserPopedomPass<>1 then
if DateDiff("s",Request.Cookies("posttime"),Now()) < int(PostTime) then message=message&"<li>为防止有人用程序灌水，论坛限制一个人两次发帖间隔必须大于"&PostTime&"秒！"
end if

if message<>"" then error(""&message&"")

sql="select * from [user] where username='"&HTMLEncode(username)&"'"
rs.Open sql,Conn,1,3


if rs("userlife")<3 then
message=message&"<li>您的体力值 < <FONT color=red>3</FONT> 不能发表文章<li>每有效停留时间<FONT color=red> 10 </FONT>分钟：体力值：<FONT color=red>+10</FONT>"
end if

StopPostTime=int(DateDiff("s",rs("regtime"),Now())/60)
if StopPostTime < int(Reg10) then message=message&"<li>新注册用户必须 "&Reg10&" 分钟后才能发帖！"

if message<>"" then error(""&message&"")


if UserPopedomPass<>1 then rs("userlife")=rs("userlife")-1
rs("postrevert")=rs("postrevert")+1
rs("money")=rs("money")+1
rs("experience")=rs("experience")+1
rs("landtime")=now()
rs.update
rs.close



sql="insert into [reforum"&ReList&"] (topicid,username,content,postip) values ('"&id&"','"&Request.Cookies("username")&"','"&content&"','"&Request.ServerVariables("REMOTE_ADDR")&"')"
conn.Execute(SQL)

conn.execute("update [forum] set lastname='"&username&"',replies=replies+1,lasttime='"&now()&"' where ID="&ID&"")
conn.execute("update [bbsconfig] set lasttopic='<a href=ShowPost.asp?id="&id&">"&left(topic,15)&"</a>',lastname='"&username&"',lasttime='"&now()&"',today=today+1,tolrestore=tolrestore+1 where id="&forumid&"")
conn.execute("update [clubconfig] set today=today+1")



Response.Cookies("posttime")=now

Application.Lock
Application(CacheName&"CountReforum") = Application(CacheName&"CountReforum")+1
Application.UnLock

message=message&"<li>回复主题成功<li><a href=ShowPost.asp?id="&id&">返回主题</a><li><a href=ShowForum.asp?forumid="&forumid&">返回论坛</a><li><a href=Default.asp>返回论坛首页</a>"
'奖励
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update 用户 set 泡豆点数=泡豆点数+"&bbs_add2&" where 姓名='"&jhname&"'")
connjh.close
set connjh=nothing
'奖励
mess="<font color=000099><b>〖回复主题〗</b></font><font color=#cc0000>"&jhname&"</font><font color=006600>在江湖论坛回复了贴子</font><a href=../bbs/ShowPost.asp?id="&id&" target=_blank title=点击便可查看该主题><b>["&topic&"]</b></a>[站长特奖励"&bbs_add2&"个豆点]"

'发贴或回贴在聊天室显示
dim says,act,towhoway,towho,addwordcolor,saycolor,addsays,saystr,i,mess
if mess<>"" then
says="<bgsound src=wav/xintie.wav loop=1><img src=img/xintie.gif>"&mess
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & jhname & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
addmsg saystr
Function Yushu(a)
 Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
end if
'发贴或回贴在聊天室显示
succeed(""&message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?forumid="&forumid&">")


end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






if Request("quote")=1 and Request("preview")="" then

if isnumeric(""&Request("retopicid")&"") then
sql="select * from [reforum"&ReList&"] where id="&Request("retopicid")&""
else
sql="select * from [forum] where ID="&id&""
end if
Set Rs=Conn.Execute(sql)
content =rs("content")
quote="<blockquote><strong>引用</strong>：<hr>原文由 <b>"&rs("username")&"</b> 发表于 <i>"&rs("posttime")&"</i> :<br>"&content&""&vbCrlf&"<hr></blockquote>"
rs.close

end if


%>
<script>
if ("<%=logo%>"!=''){logo.innerHTML="<img border=0 src=<%=logo%> onload='javascript:if(this.height>60)this.height=60;'>"}
</script>

	<table border="0" width="97%" align="center" cellspacing="1" cellpadding="4" class=a2>
		<tr class=a3>
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <%ForumTree(followid)%><%=ForumTreeList%> <a href=ShowForum.asp?forumid=<%=forumid%>><%=bbsname%></a> → <a href="ShowPost.asp?id=<%=id%>"><%=topic%></a> → 回复帖子</td>
		</tr>
	</table><br>



<SCRIPT>valigntop()</SCRIPT>
<TABLE cellSpacing=1 cellPadding=5 width=97% border=0 class=a2 align="center">





<form name="yuziform" method="post" onSubmit="return CheckForm(this);">
<input name="content" type="hidden" value='<%=quote%>'>
<input type="hidden" value="<%=Request("topic")%>" name="topic">

<input type=hidden name=id value=<%=id%>>

<TR>
<TD vAlign=left class=a1 colSpan=2 height=25><b>回复文章</b></TD></TR>
<TR>
<TD width=160 class=a4 height=25><B>用户名称</B></TD>
<TD width=574 class=a4 height=25><%=Request.Cookies("username")%> [<a href="login.asp?menu=out">退出登录</a>]</TD>
</TR>
<TR>
<TD vAlign=top class=a3>
<TABLE cellSpacing=0 cellPadding=0 width=100% align=left border=0>

<TR>
<TD vAlign=top align=left width=100% class=a3><B>文章内容</B><BR>
(<a href="javascript:CheckLength();">查看内容长度</a>)<BR><BR>
</TD></TR>
<TR>
<TD vAlign=center align=center width=100%>
<TABLE
style="BORDER-RIGHT: 1px inset; BORDER-TOP: 1px inset; BORDER-LEFT: 1px inset; BORDER-BOTTOM: 1px inset"
cellSpacing=1 cellPadding=3 border=0>
<TR align=middle>
<script>
ii=0
for(i=1;i<=25;i++) {
index = Math.floor(Math.random() * 80+1);
ii=ii+1
document.write("<TD><a href=javascript:emoticon('"+index+"')><img border=0 src=images/Emotions/"+index+".gif></a></TD>")
if (ii==5){document.write("</TR><TR align=middle>");ii=0;}
}
</script>

</TR></TABLE>
</TD></TR>
</TABLE></TD>
<TD class=a3 height=250>

<SCRIPT src="inc/post.js"></SCRIPT>
</TD></TR>



<TR>
<TD align=left class=a4>
<IMG src=images/affix.gif><b>增加附件</b></TD>

<TD align=left class=a4><font color="FFFFFF"><b><IFRAME src="upfile.asp" frameBorder=0 width="100%" scrolling=no height=21></IFRAME></b></font></TD></TR>

<TR>
<TD align=middle class=a3 colSpan=2 height=27>
<INPUT tabIndex=4 type=submit value=回复主题 name=submit1>&nbsp; <input name=preview type="Button" value=" 预 览 " onclick="Gopreview()">   <INPUT type=reset value=" 重 置 "></TD></TR></FORM>
</TABLE>
<SCRIPT>valignbottom()</SCRIPT>
<form name=preview action=preview.asp method=post target=preview_page><input name="content" type="hidden"></form>

<%

htmlend
%>