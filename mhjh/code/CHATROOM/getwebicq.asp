<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
webicq=Application("yx8_mhjh_webicq")
webicqnum=Application("yx8_mhjh_webicqnum")
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
Application("yx8_mhjh_webicq")=index
else
Application("yx8_mhjh_webicq")=newwebicq
end if
Application("yx8_mhjh_webicqname")=newwebicqname
Application.UnLock
erase newwebicq
erase webicq
%>
<head>
<bgsound src='../mid/dh.wav' loop=1>
<link rel=stylesheet href='css3.css'>
</head>
<body background='bg1.gif'>
<div align=center>
<font color=0000FF>飞鸽传书</font><br>
<table width='100%'>
<%=msg%>
</table>
</div>
</body>