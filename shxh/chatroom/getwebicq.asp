<%
Response.Expires=-1
username=session("Ba_jxqy_username")
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
webicq=Application("Ba_jxqy_webicq")
webicqnum=Application("Ba_jxqy_webicqnum")
webicqubd=ubound(webicq)
dim newwebicq()
newwebicqname=";"
j=1
for i=1 to webicqubd step 5
	if datediff("n",webicq(i+4),nowtime)<10 then
		if webicq(i+2)<>username then
			redim preserve newwebicq(j),newwebicq(j+1),newwebicq(j+2),newwebicq(j+3),newwebicq(j+4)
			newwebicqname=newwebicqname&webicq(i+2)&";"
			newwebicq(j)=webicq(i)
			newwebicq(j+1)=webicq(i+1)
			newwebicq(j+2)=webicq(i+2)
			newwebicq(j+3)=webicq(i+3)
			newwebicq(j+4)=webicq(i+4)
			j=j+5
		elseif webicq(i+2)=username	then
			msg=msg+"<tr><td bgcolor=00FF00 width='100%'>来自：<font color=ff0000>"&webicq(i+1)&"</font></td></td><tr><td width=140>"&webicq(i+3)&"</td></tr>"
		end if
	end	if
next
Application.Lock
if j=1 then
	dim index(0)
	Application("Ba_jxqy_webicq")=index
else	
	Application("Ba_jxqy_webicq")=newwebicq
end if	
Application("Ba_jxqy_webicqname")=newwebicqname
Application.UnLock
erase newwebicq
erase webicq
%>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' >
<div align=center>
<font color=0000FF>飞鸽传书</font><br>
<table width='100%'>
<%=msg%><bgsound src='../mid/dh.wav' loop=1>
</table>
</div>
</body>