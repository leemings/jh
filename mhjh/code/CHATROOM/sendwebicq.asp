<%
sendto=Request.Form("sendto")
message=trim(Request.Form("message"))
message=server.HTMLEncode(message)
username=session("yx8_mhjh_username")
webicq=Application("yx8_mhjh_webicq")
webicqnum=Application("yx8_mhjh_webicqnum")
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
newwebicq(j+3)=message
newwebicq(j+4)=nowtime
newwebicqname=newwebicqname&sendto&";"
Application.Lock
Application("yx8_mhjh_webicqnum")=webicqnum+1
Application("yx8_mhjh_webicq")=newwebicq
Application("yx8_mhjh_webicqname")=newwebicqname
Application.UnLock
erase newwebicq
erase webicq
%>
<head>
<link rel='stylesheet' href='css2.css'>
<title></title>
</head>
<body  oncontextmenu=self.event.returnValue=false background="bg1.gif">
<div align=center>
<font color=0000FF>飞鸽传书</font><br>
让我们一起祈祷吧，让光荣与梦想随信鸽一起苍茫远去！
</div>
</body>
