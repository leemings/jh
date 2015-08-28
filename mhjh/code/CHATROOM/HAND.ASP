<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Session("yx8_mhjh_userchatroomsn")
if application("yx8_mhjh_klt")<>0 then
application("yx8_mhjh_klt")=0
response.write "你刚才放的僵尸没有打完，请等一下再放"
response.end
else
randomize()
rnd1=cint(rnd()*350)
jstl=int(rnd*50000)+100
		Application.Lock
		Application("yx8_mhjh_klt")=jstl
		Application.UnLock	
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
		newtalkarr(594)=username
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>【僵尸】僵尸呀有僵尸呀！一女子尖叫，一头僵尸闯进聊天室，僵尸杀伤力力："&jstl&"</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=klt.asp?tl="&Application("yx8_mhjh_klt")&" target=optfrm><img src=data/kl.GIF border=0></a></marquee>"
        newtalkarr(600)=chatroomsn
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr		
end if
response.write "放出僵尸成功"
%>