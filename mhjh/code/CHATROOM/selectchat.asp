<%
Response.Expires=-1
chatroomnum=Application("yx8_mhjh_chatroomnum")
chatroomsn=session("yx8_mhjh_userchatroomsn")
msg="<head><script language=javascript>function chgchat(){chatsn=document.form1.sele1.value;parent.sendfrm.location.href='about:blank';parent.getfrm.location.href='about:blank';parent.resfrm.location.href='chgchat.asp?chatroomsn='+chatsn;}function init(){var chatsn='"&chatroomsn&"';var chatnum='"&chatroomnum&"';var i;for(i=0;i<chatnum;i++){if(document.form1.sele1.options[i].value==chatsn){document.form1.sele1.options[i].selected=true;}}}</script></head><body background='bg1.gif' onload='javascript:init();' oncontextmenu=self.event.returnValue=false><div align=center><form name=form1><select name=sele1 onchange='javascript:chgchat();'>"
for i=1 to chatroomnum
msg=msg&"<option value='"&i&"'>"&Application("yx8_mhjh_systemname"&i)&"</option>"
next
msg=msg&"</select></form></div></body>"
Response.Write msg
%>
