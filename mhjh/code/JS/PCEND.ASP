<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("yx8_mhjh_username")
lastip=Request.ServerVariables("LOCAL_ADDR")
templastip=StrReverse(left(lastip&"9fhb6utime,./",15))
templenlastip=len(lastip)
mmcheck=""
for j=1 to 15
mmcheck=mmcheck+chr(asc(mid(templastip,j,1))-templenlastip+int(j*1.1))
next
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from pc21 where username='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
zpoint=rst("zpoint")
upoint=rst("upoint")
bet=rst("bet")
msg="庄家"&zpoint&"点，你"&upoint&"点，"
if not(upoint>21 or zpoint>=upoint) then
conn.Execute "update 用户 set 银两=银两+"&bet*2&" where 姓名='"&username&"'"
conn.Execute "update pc21 set win=win+1,zpoint=21 where username='"&username&"'"
msg=msg&"你赢了"&bet&"两银子。"
	kl="<font color=0000FF>"&username&"</font>玩经典21点发财了!!赢了"&bet&"两银子,够帅气！"
else
conn.Execute "update pc21 set zpoint=21,bet=0 where username='"&username&"'"
	kl="<font color=0000FF>"&username&"</font>玩经典21点,输得很惨,"&bet&"两白银打了水漂！"
msg=msg&"不好意思，你输了"&bet&"两银子。"
end if
msg=msg&"<script language=javascript>parent.imgfrm.showcard("&rst("z0")&","&rst("z1")&","&rst("z2")&","&rst("z3")&","&rst("z4")&","&rst("z5")&","&rst("u0")&","&rst("u1")&","&rst("u2")&","&rst("u3")&","&rst("u4")&","&rst("u5")&");</script>"
says="<font color=red><b>【娱乐赌博】</b></font>"&kl	
dim newtalkarr(600) 
talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=name
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="008000" 
		newtalkarr(599)=""&says&""
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
%>
<html>
<head>
<link rel=stylesheet href='../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'>
<p align=center>
<%=msg%>
<input type=button value='再来一局' onclick=javascript:location.replace('pcbet.asp');>
</p>
</body>
</html>
