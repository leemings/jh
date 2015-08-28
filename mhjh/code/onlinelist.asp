<%
Response.Expires=-1
randomize timer
s=1+int(rnd*4)
chatroomnum=Application("yx8_mhjh_chatroomnum")
msg="<body topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0><select name=D6 ><option value='#' style='background:FFFFFF'>"&Application("yx8_mhjh_allonlinenum")+s+3&"人在线/"&Application("yx8_mhjh_allonlinenum")+3&"人聊天</option>"
onlinelist=Application("yx8_mhjh_onlinelist")
onlinelistubd=ubound(onlinelist)
for j=1 to onlinelistubd step 6
if onlinelist(j)=Application("yx8_mhjh_admin") then 
if Application("yx8_mhjh_yinshen")=True then 
        msg=msg
else
        msg=msg&"<option value='#'>"&onlinelist(j)&"</option>"
end if
else
	msg=msg&"<option value='#'>"&onlinelist(j)&"</option>"
end if
next
msg=msg&"</select></body>"
%>
document.write("<%=msg%>");



