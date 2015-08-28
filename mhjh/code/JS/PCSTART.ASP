<%
pcbet=Request.Form("pcbet")
if not isnumeric(pcbet) then Response.Redirect "../error.asp?id=024"
pcbet=clng(pcbet)
if pcbet>100001 or pcbet<1 then Response.Redirect "../error.asp?id=024"
username=Session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
dim zpcarr(5)
zpoint=0
randomize()
for i=0 to 5
rndz=clng(rnd()*51+1)
zpointadd=rndz mod 13
if zpointadd=0 or zpointadd>10 then zpointadd=1
if zpoint+zpointadd>21 then
for j=i to 5
zpcarr(j)=53
next
exit for
else
zpoint=zpoint+zpointadd
zpcarr(i)=rndz
end if
next
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&username&"' and 银两>="&pcbet,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=030"
rst.Close
rst.Open "select * from pc21 where username='"&username&"'",conn
if rst.EOF or rst.BOF then
conn.Execute "insert into pc21 values('"&username&"',"&zpcarr(0)&","&zpcarr(1)&","&zpcarr(2)&","&zpcarr(3)&","&zpcarr(4)&","&zpcarr(5)&","&zpoint&",53,53,53,53,53,53,0,"&pcbet&",0)"
else
conn.Execute "update pc21 set z0="&zpcarr(0)&",z1="&zpcarr(1)&",z2="&zpcarr(2)&",z3="&zpcarr(3)&",z4="&zpcarr(4)&",z5="&zpcarr(5)&",zpoint="&zpoint&",u0=53,u1=53,u2=53,u3=53,u4=53,u5=53,upoint=0,bet="&pcbet&" where username='"&username&"'"
end if
rst.Close
set rst=nothing
conn.Execute "update 用户 set 银两=银两-"&pcbet&" where 姓名='"&username&"'"
	kl="<font color=0000FF>"&username&"</font>相中了经典21点,希望通过赌博发财,一次投入"&pcbet&"两白银,现在准备开牌,大家祝他好运！"
conn.Close
set conn=nothing
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
response.write "<html><head><link rel=stylesheet href='../style.css'></head><body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'><p align=center><input type=button value=' 开始发牌 ' onclick=javascript:location.replace('pccontinue.asp');></p><script language=javascript>parent.imgfrm.showcard(53,53,53,53,53,53,53,53,53,53,53,53);</script></body></html>"
%>
