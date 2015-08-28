<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
usercorp=session("yx8_mhjh_usercorp")
if username="" then Response.Redirect "../error.asp?id=016"
if username<>Application("yx8_mhjh_admin") or usercorp<>"官府" then Response.Redirect "../error.asp?id=046"
if Application("yx8_mhjh_chatroomsnkb")=1 then
   Application("yx8_mhjh_chatroomsnkb")=0   
else
   Application("yx8_mhjh_chatroomsnkb")=1
end if
%>
<head>
<title></title>
</head>
<body oncontextmenu=self.event.returnValue=false background="bg1.gif" topmargin="100" >
成功将聊天大厅设置为:
<%if Application("yx8_mhjh_chatroomsnkb")=1 then%>
开打
<%else%>
关打
<%end if%>

