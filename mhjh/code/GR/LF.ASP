<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then 
Response.Redirect "../error.asp?id=016"
lastip=Request.ServerVariables("LOCAL_ADDR")
templastip=StrReverse(left(lastip&"9fhb6utime,./",15))
templenlastip=len(lastip)
mmcheck=""
for j=1 to 15
mmcheck=mmcheck+chr(asc(mid(templastip,j,1))-templenlastip+int(j*1.1))
next


else
%>
<HTML>
<HEAD>
<title><%=Application("yx8_mhjh_systemname")%></title>
<link rel="stylesheet" href="../style.css">
</HEAD>
<body  oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'>
<p align=center> <b> <font color="#000000" size="4">我要逃婚</font></b></p>

<p align="center"><img border="0" src="../image/no_27.gif"></p>
<p align="center">我是超时空博士，我现在是回到这个时代给你们解答</p>
<p align="center">问题，如果你不想要现在的配偶，或者你很少见你的</p>
<p align="center">配偶，觉得寂寞难耐了。那么，我可以帮你<a href="left.asp"><font color="#0000FF">实现逃婚</font></a></p>
<p align="center">　</p>
<p>　</p>

</body></HTML> 
<%
end if
%>