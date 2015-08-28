<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if session("yx8_mhjh_username")="" then Response.Redirect "../error.asp?id=016"
if session("yx8_mhjh_username")<>Application("yx8_mhjh_admin") then Response.Redirect "../error.asp?id=016"
if Application("yx8_mhjh_yinshen")=true then 
Application("yx8_mhjh_yinshen")=false
else
Application("yx8_mhjh_yinshen")=true
end if
%>
<head>
<title></title>
</head>
<body oncontextmenu=self.event.returnValue=false background="../chatroom/bg1.gif" topmargin="100" >
成功将隐身登陆状态设置为:
<%if Application("yx8_mhjh_yinshen")=true then%>
隐身
<%else%>
现身
<%end if%>

