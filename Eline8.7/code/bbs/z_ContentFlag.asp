<li>HTML标签：<%=iif(Board_Setting(5),"<font color=" & Forum_body(8) & "><b>√</b></font>","<font color=" & Forum_body(8) & "><b>×</b></font>")%>
<li>UBB标签：<%=iif(Board_Setting(6),"<font color=" & Forum_body(8) & "><b>√</b></font>","<font color=" & Forum_body(8) & "><b>×</b></font>")%>
<li>帖图标签：<%=iif(Board_Setting(7),"<font color=" & Forum_body(8) & "><b>√</b></font>","<font color=" & Forum_body(8) & "><b>×</b></font>")%>
<li>多媒体标签：<%=iif(Board_Setting(9),"<font color=" & Forum_body(8) & "><b>√</b></font>","<font color=" & Forum_body(8) & "><b>×</b></font>")%>
<li>表情字符转换：<%=iif(Board_Setting(8),"<font color=" & Forum_body(8) & "><b>√</b></font>","<font color=" & Forum_body(8) & "><b>×</b></font>")%>
<li>上传文件：<%=iif(GroupSetting(7)=1 or GroupSetting(7)=2,"<font color=" & Forum_body(8) & "><b>√</b></font>","<font color=" & Forum_body(8) & "><b>×</b></font>")%>
<li>最多：<font color=<%=Forum_body(8)%>><b><%=Board_Setting(16)\1024%>KB</b></font>
<%select case cint(forum_setting(73))
case 1%>
	<li>今日还可发表主题：<font color=blue><b><%=myUserPPD*cint(forum_setting(74))-myTodayTopic%></b></font>个
<%case 2%>
	<li>今日还可发表回复：<font color=blue><b><%=myUserPPD*cint(forum_setting(74))-myTodayReply%></b></font>个
<%case 3%>
	<li>今日还可发表帖子：<font color=blue><b><%=myUserPPD*cint(forum_setting(74))-myTodayTopic-myTodayReply%></b></font>个
<%end select%>
<li>发帖机遇：<%=iif(Board_Setting(44),"<font color=" & Forum_body(8) & "><b>√</b></font>","<font color=" & Forum_body(8) & "><b>×</b></font>")%>
