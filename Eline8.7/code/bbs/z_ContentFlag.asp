<li>HTML��ǩ��<%=iif(Board_Setting(5),"<font color=" & Forum_body(8) & "><b>��</b></font>","<font color=" & Forum_body(8) & "><b>��</b></font>")%>
<li>UBB��ǩ��<%=iif(Board_Setting(6),"<font color=" & Forum_body(8) & "><b>��</b></font>","<font color=" & Forum_body(8) & "><b>��</b></font>")%>
<li>��ͼ��ǩ��<%=iif(Board_Setting(7),"<font color=" & Forum_body(8) & "><b>��</b></font>","<font color=" & Forum_body(8) & "><b>��</b></font>")%>
<li>��ý���ǩ��<%=iif(Board_Setting(9),"<font color=" & Forum_body(8) & "><b>��</b></font>","<font color=" & Forum_body(8) & "><b>��</b></font>")%>
<li>�����ַ�ת����<%=iif(Board_Setting(8),"<font color=" & Forum_body(8) & "><b>��</b></font>","<font color=" & Forum_body(8) & "><b>��</b></font>")%>
<li>�ϴ��ļ���<%=iif(GroupSetting(7)=1 or GroupSetting(7)=2,"<font color=" & Forum_body(8) & "><b>��</b></font>","<font color=" & Forum_body(8) & "><b>��</b></font>")%>
<li>��ࣺ<font color=<%=Forum_body(8)%>><b><%=Board_Setting(16)\1024%>KB</b></font>
<%select case cint(forum_setting(73))
case 1%>
	<li>���ջ��ɷ������⣺<font color=blue><b><%=myUserPPD*cint(forum_setting(74))-myTodayTopic%></b></font>��
<%case 2%>
	<li>���ջ��ɷ���ظ���<font color=blue><b><%=myUserPPD*cint(forum_setting(74))-myTodayReply%></b></font>��
<%case 3%>
	<li>���ջ��ɷ������ӣ�<font color=blue><b><%=myUserPPD*cint(forum_setting(74))-myTodayTopic-myTodayReply%></b></font>��
<%end select%>
<li>����������<%=iif(Board_Setting(44),"<font color=" & Forum_body(8) & "><b>��</b></font>","<font color=" & Forum_body(8) & "><b>��</b></font>")%>
