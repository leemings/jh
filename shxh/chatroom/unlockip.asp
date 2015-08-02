<%
Response.Expires=-1
chatroomsn=Session("Ba_jxqy_userchatroomsn")
mycorp=session("Ba_jxqy_usercorp")
mygrade=session("Ba_jxqy_usergrade")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
Response.Write bgcolor
chatroomname=Application("Ba_jxqy_systemname"&chatroomsn)
lockiplist=Application("Ba_jxqy_lockip"&chatroomsn)
lockip=split(lockiplist,";")
lockipubd=ubound(lockip)
Response.Write "<head><style type='text/css'>body{font-size:12pt}a{text-decoration:none;color:0000FF; font-size: 12pt; font-family: 幼圆}a:visited{text-decoration:none;color:0000FF; font-size: 12pt; font-family: 幼圆}</style></head><body  bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu=self.event.returnValue=false><div align=center><font size=5 color=0000ff>"&chatroomname&"</font><br><font color=FF0000>"&Application("Ba_jxqy_onlinenum"&chatroomsn)&"</font>/<font color=008800>"&Application("Ba_jxqy_allonlinenum")&"</font>人在线<hr><table width='100%'>"
if (mygrade>unlockipright and mycorp="官府") or (chatroomname=mycorp and mygrade>8) then 
for i=1 to lockipubd-1
	Response.Write "<tr><td><a href=javascript:parent.talkfrm.settalk('//解锁','"&lockip(i)&"') target=talkfrm onmouseover=""window.status='官府中"&unlockipright&"级以上人员适用，或本派9级以上长老也可以在本门内部使用';return true;"" onmouseout=""window.status='';return true;"" title='官府中"&unlockipright&"级以上人员适用，或本派9级以上长老也可以在本门内部使用'>"&lockip(i)&"</a></td></tr>"	
next
end if
Response.Write "</table></body>"
%>