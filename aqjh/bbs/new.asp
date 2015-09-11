<!-- #include file="conn.asp" -->
<%
'=========================================================
' 调用标题代码
'<span id=yuzi>Loading...</span>
'<script src=new.asp?id=yuzi&forumid=1&TopicCount=10&Count=15&showtime=1&icon=◎></script>
'---------------------------------------------------------
' 调用内容代码
'<span id=yuzi>Loading...</span>
'<script src=new.asp?id=yuzi&showcontent=1&icon=◎></script>
'---------------------------------------------------------
' id:         定义识别代码[其中“yuzi”可自定义，如一个页面多个调用请各自更改不同的代码，共2处需要修改]
' forumid:    指定论坛的ID
' TopicCount: 显示多少条主题
' Count:      标题 长度
' showtime:   1=显示时间 0=不显示时间
' showcontent:1=显示内容 0=不显示内容
' icon:       标题前的符号
'=========================================================
forumid=int(Request("forumid"))
if Request("TopicCount")=empty then
TopicCount=10
else
TopicCount=int(Request("TopicCount"))
end if

if Request("Count")=empty then
Count=100
else
Count=int(Request("Count"))
end if



cluburl=Conn.Execute("Select cluburl From clubconfig")(0)

	if forumid <> empty then sqlforumid=" and forumid="&forumid&""
	sql="select top "&TopicCount&" * from forum where deltopic=0 "&sqlforumid&" order by id desc"
	Set Rs=Conn.Execute(sql)
	do while Not RS.Eof
	if Request("showcontent")=1 then
	list=list&" <table border=0 width="&CHR(34)&"100%"&CHR(34)&" style=TABLE-LAYOUT:fixed><tr><td>"&Request("icon")&" <a href="&cluburl&"ShowPost.asp?id="&rs("id")&" target=_blank>"&rs("Topic")&"</a></td></tr><tr><td height=100>"&rs("content")&"</td></tr><tr><td align=right><a href="&cluburl&"ShowPost.asp?id="&rs("id")&" target=_blank><u> "&rs("username")&" 发表于 "&rs("posttime")&" - "&rs("replies")&" 个评论</u></a></td></tr></table><hr>"
	else
set reg=new regexp'过滤HTML标签
reg.Global=true
reg.Pattern="<[^>]*>"
topic=reg.Replace(rs("topic"),"")
set reg=nothing
	if Request("showtime")=1 then showtime=" ["&rs("lasttime")&"]"
	list=list&""&Request("icon")&" <a href="&cluburl&"ShowPost.asp?id="&rs("id")&" target=_blank>"&left(topic,Count)&"</a>"&showtime&"<br>"
	end if


	RS.MoveNext
	Loop
	if list="" then list="·<a href=bbs/ target=_blank>暂无新贴,请登录论坛!</a>"
	set rs=nothing
	conn.close
	set conn=nothing
%>
<%=Request("id")%>.innerHTML='<%=list%>'
