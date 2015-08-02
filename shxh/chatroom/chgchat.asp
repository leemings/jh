<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
dim index(0)
newchatroomsn=Request.QueryString("chatroomsn")
chatroomsn=session("Ba_jxqy_userchatroomsn")
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "chaterror.asp?id=004"
if newchatroomsn="" then 
	newchatroomsn=1
elseif not isnumeric(newchatroomsn) then
	newchatroomsn=1
else
	newchatroomsn=cint(newchatroomsn)	
end if
if newchatroomsn<1 or newchatroomsn>Application("Ba_jxqy_chatroomnum") then newchatroomsn=1
newchatroomname=Application("Ba_jxqy_systemname"&newchatroomsn)
chatroomname=Application("Ba_jxqy_systemname"&chatroomsn)
onlinelist=Application("Ba_jxqy_onlinelist")
chatroomnum=Application("Ba_jxqy_chatroomnum")
maxnosaytime=Application("Ba_jxqy_nosaytime")
onlinelistubd=ubound(onlinelist)
nowtime=now()
msg=replace(Application("Ba_jxqy_guestleft"),"@#",Application("Ba_jxqy_systemname"&chatroomsn))&"，可能是要到"&Application("Ba_jxqy_systemname"&newchatroomsn)&"去吧。"
msg2=replace(Application("Ba_jxqy_guestjoin"),"@#",Application("Ba_jxqy_systemname"&newchatroomsn))
j=1
dim newonlinelist()
dim newonlinename()
for i=1 to chatroomnum
	redim preserve newonlinename(i)
	newonlinename(i)=";"
next
newonlinenum=0
for i=1 to onlinelistubd step 8
	if datediff("s",onlinelist(i+4),nowtime)<=maxnosaytime and onlinelist(i)=username then
		newonlinenum=newonlinenum+1
		newonlinename(newchatroomsn)=newonlinename(newchatroomsn)&username&";"
		redim preserve newonlinelist(j),newonlinelist(j+1),newonlinelist(j+2),newonlinelist(j+3),newonlinelist(j+4),newonlinelist(j+5),newonlinelist(j+6),newonlinelist(j+7)
		newonlinelist(j)=onlinelist(i)
		newonlinelist(j+1)=onlinelist(i+1)
		newonlinelist(j+2)=onlinelist(i+2)
		newonlinelist(j+3)=onlinelist(i+3)
		newonlinelist(j+4)=onlinelist(i+4)
		newonlinelist(i+5)=newchatroomsn
		newonlinelist(j+6)=onlinelist(i+6)
		newonlinelist(j+7)=onlinelist(i+7)
		j=j+8
	elseif datediff("s",onlinelist(i+4),nowtime)<=maxnosaytime and onlinelist(i)<>username then
		newonlinenum=newonlinenum+1
		newonlinename(onlinelist(i+5))=newonlinename(onlinelist(i+5))&onlinelist(i)&";"
		redim preserve newonlinelist(j),newonlinelist(j+1),newonlinelist(j+2),newonlinelist(j+3),newonlinelist(j+4),newonlinelist(j+5),newonlinelist(j+6),newonlinelist(j+7)
		newonlinelist(j)=onlinelist(i)
		newonlinelist(j+1)=onlinelist(i+1)
		newonlinelist(j+2)=onlinelist(i+2)
		newonlinelist(j+3)=onlinelist(i+3)
		newonlinelist(j+4)=onlinelist(i+4)
		newonlinelist(j+5)=onlinelist(i+5)
		newonlinelist(j+6)=onlinelist(i+6)
		newonlinelist(j+7)=onlinelist(i+7)
		j=j+8
	else
		Application.Lock
		Application("Ba_jxqy_allonline")=replace(Application("Ba_jxqy_allonline"),";"&onlinelist(i)&";",";")
		Application.UnLock
	end if
next
Application.Lock
if newonlinenum=0 then
	Application("Ba_jxqy_onlinelist")=index
else
	Application("Ba_jxqy_onlinelist")=newonlinelist
end if
Application("Ba_jxqy_allonlinenum")=newonlinenum
for i=1 to chatroomnum
	Application("Ba_jxqy_onlinename"&i)=newonlinename(i)
next
Application.UnLock
talkarr=Application("Ba_jxqy_talkarr")
talkpoint=Application("Ba_jxqy_talkpoint")
dim newtalkarr(600)
j=1
for i=21 to 600 step 10
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
newtalkarr(581)=talkpoint+1
newtalkarr(582)=1
newtalkarr(583)=0
newtalkarr(584)=username
newtalkarr(585)="大家"
newtalkarr(586)=""
newtalkarr(587)="000000"
newtalkarr(588)="000000"
newtalkarr(589)="<font color=FF0000>【公告】</font>"&msg&"<font class=timsty>"&time()&"</font>"
newtalkarr(590)=chatroomsn
newtalkarr(591)=talkpoint+1
newtalkarr(592)=1
newtalkarr(593)=0
newtalkarr(594)=username
newtalkarr(595)="大家"
newtalkarr(596)=""
newtalkarr(597)="000000"
newtalkarr(598)="000000"
newtalkarr(599)="<font color=FF0000>【公告】</font>"&msg2&"<font class=timsty>"&time()&"</font>"
newtalkarr(600)=newchatroomsn
Application.Lock
Application("Ba_jxqy_talkpoint")=talkpoint+1
Application("Ba_jxqy_talkarr")=newtalkarr
Application.UnLock
session("Ba_jxqy_userchatroomsn")=newchatroomsn
%>
<script language=javascript>
location.replace('onlinelist.asp');
parent.getfrm.location.href='getmsg.asp';
</script>