<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
chatroomnum=Application("Ba_jxqy_chatroomnum")
msg="<body topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0><select name=D6 ><option value='#' style='background:FFFF00'>--共"&Application("Ba_jxqy_allonlinenum")&"人在线--</option>"
onlinelist=Application("Ba_jxqy_onlinelist")
onlinelistubd=ubound(onlinelist)
for j=1 to onlinelistubd step 8
	msg=msg&"<option value='#'>"&onlinelist(j)&"</option>"
next
msg=msg&"</select></body>"
%>
document.write("<%=msg%>");



