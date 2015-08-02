<%
sendto=Request.Form("sendto")
message=trim(Request.Form("message"))
message=server.HTMLEncode(message)
lasttalktime=session("mylasttalktime")
username=session("Ba_jxqy_username")
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
nowtime=now()
if datediff("s",lasttalktime,nowtime)<10 then
Response.Write "<script language=javascript>alert('对不起，要点歌吗吗？请等十秒钟！');history.back();</script>"
Response.End
end if
webicq=Application("Ba_jxqy_webicq")
webicqnum=Application("Ba_jxqy_webicqnum")
webicqubd=ubound(webicq)
dim newwebicq()
newwebicqname=";"
j=1
for i=1 to webicqubd step 5
	if datediff("n",webicq(i+4),nowtime)<10 then
		redim preserve newwebicq(j),newwebicq(j+1),newwebicq(j+2),newwebicq(j+3),newwebicq(j+4)
		newwebicqname=newwebicqname&webicq(i+2)&";"
		newwebicq(j)=webicq(i)
		newwebicq(j+1)=webicq(i+1)
		newwebicq(j+2)=webicq(i+2)
		newwebicq(j+3)=webicq(i+3)
		newwebicq(j+4)=webicq(i+4)
		j=j+5
	end if
next
redim preserve newwebicq(j),newwebicq(j+1),newwebicq(j+2),newwebicq(j+3),newwebicq(j+4)
newwebicq(j)=webicqnum+1
newwebicq(j+1)=username
newwebicq(j+2)=sendto
newwebicq(j+3)=message&Request.Form("song")
newwebicq(j+4)=nowtime
newwebicqname=newwebicqname&sendto&";"
Application.Lock
Application("Ba_jxqy_webicqnum")=webicqnum+1
Application("Ba_jxqy_webicq")=newwebicq
Application("Ba_jxqy_webicqname")=newwebicqname
Application.UnLock
erase newwebicq
erase webicq
%>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' >
<div align=center><font size=5 color=0000ff><%=Application("Ba_jxqy_systemname"&chatroomsn)%></font><br><font color=FF0000><%=Application("Ba_jxqy_onlinenum"&chatroomsn)%></font>/<font color=008800><%=Application("Ba_jxqy_allonlinenum")%></font>人在线<hr>
<font color=0000FF>好歌送好友</font><br><%=message%>
</div>
</body>
