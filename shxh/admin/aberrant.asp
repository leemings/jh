<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Request.QueryString("chatroomsn")
opt=Request.QueryString("opt")
st=Request.QueryString("touser")
chatroomname=Application("Ba_jxqy_systemname"&chatroomsn)
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
arrestright=Application("Ba_jxqy_arrestright")
gaolright=Application("Ba_jxqy_gaolright")
lockipright=Application("Ba_jxqy_lockipright")
unlockipright=Application("Ba_jxqy_unlockipright")
bombright=application("Ba_jxqy_bombright")
exaltgraderight=Application("Ba_jxqy_exaltgraderight")
declinegraderight=Application("Ba_jxqy_declinegraderight")
chatroomnum=Application("Ba_jxqy_chatroomnum")
un=session("Ba_jxqy_username")
if un="" then Response.redirect "error.asp?id=016"
tt=now()
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&un&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
ungrade=rst("等级")
uncorp=rst("门派")
unidentity=rst("身份")
rst.Close
rst.Open "select * from 用户 where 姓名='"&st&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
stgrade=rst("等级")
rst.Close
set rst=nothing
if ((ungrade>10 and uncorp="官府") or unidentity="掌门") and opt="驱逐" and ungrade>stgrade then
	lt=dateadd("n",5,tt)
	msg="<font color=FF0000>【驱逐】<\/font>##将%%无情的驱逐出了"&Application("Ba_jxqy_systemname")
elseif ungrade>=arrestright and uncorp="官府" and opt="逮捕" and ungrade>stgrade then
	lt=dateadd("n",10,tt)
	conn.Execute "update 用户 set 状态='逮捕',最后登录时间='"&lt&"' where 姓名='"&st&"'"
	msg="<font color=FF0000>【逮捕】<\/font>##伸出一只巨手，象拎小鸡一样将%%抓入了官府大牢。<bgsound src=\'../mid/daipu.wav\' loop=1>"
elseif ungrade>=gaolright and uncorp="官府" and opt="入狱" and ungrade>stgrade then
	lt=dateadd("n",15,tt)
	conn.Execute "update 用户 set 状态='入狱',最后登录时间='"&lt&"' where 姓名='"&st&"'"
	msg="<font color=FF0000>【入狱】<\/font>随着一声\'威武\'的吆喝，如狼似虎的##将%%带入了重犯大牢。"
elseif  ungrade>=lockipright and uncorp="官府" and opt="锁定" and ungrade>stgrade then
	lt=dateadd("n",10,tt)
	oll=Application("Ba_jxqy_onlinelist")
	ollubd=ubound(oll)
	for i=1 to ollubd step 8
		if oll(i)=st then
			lockip=oll(i+3)
			exit for
		end if
	next
	Application.Lock 
	Application("Ba_jxqy_lockip")=Application("Ba_jxqy_lockip")&lockip&";"
	Application.UnLock
	msg="<font color=FF0000>【封锁IP】<\/font>##将%%无情的驱逐出境，同时锁定IP"&lockip
elseif ungrade>=bombright and uncorp="官府" and opt="炸弹"  and ungrade>stgrade then
	lt=dateadd("n",15,tt)
	msg="<font color=FF0000>【炸弹】<\/font>##将%%恭恭敬敬地炸上了天堂，并认认真真的考虑年年祭拜的可能性。"
else
	Response.redirect "../error.asp?id=046"
end if
conn.close
set conn=nothing
aberrantlist=Application("Ba_jxqy_aberrantlist")
aberrantlistubd=ubound(aberrantlist)
dim newaberrantlist()
newaberrantname=";"
j=1
for i=1 to aberrantlistubd step 4
	if datediff("s",aberrantlist(i+3),tt)<0 then
		redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
		newaberrantlist(j)=aberrantlist(i)
		newaberrantlist(j+1)=aberrantlist(i+1)
		newaberrantlist(j+2)=aberrantlist(i+2)
		newaberrantlist(j+3)=aberrantlist(i+3)
		j=j+4
		newaberrantname=newaberrantname&aberrantlist(i)&";"
	end if
next
redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
newaberrantlist(j)=st
newaberrantlist(j+1)=un
newaberrantlist(j+2)=opt
newaberrantlist(j+3)=lt
newaberrantname=newaberrantname&st&";"
Application.Lock
Application("Ba_jxqy_aberrantlist")=newaberrantlist
Application("Ba_jxqy_aberrantname")=newaberrantname
Application.UnLock
talkarr=Application("Ba_jxqy_talkarr")
talkpoint=clng(Application("Ba_jxqy_talkpoint"))
dim newtalkarr(600)
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
newtalkarr(592)=2
newtalkarr(593)=0
newtalkarr(594)=un
newtalkarr(595)=st
newtalkarr(596)=""
newtalkarr(597)="000000"
newtalkarr(598)="000000"
newtalkarr(599)=msg&"<font class=\'timsty\'>（"&time()&"）</font>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("Ba_jxqy_talkpoint")=talkpoint+1
Application("Ba_jxqy_talkarr")=newtalkarr
Application.UnLock
Response.write "<head><link rel='stylesheet' href='../chatroom/style1.css'>.</head><body oncontextmenu=self.event.returnValue=false background='"&bgimage&"' bgcolor='"&bgcolor&"'><div heigh=100% align=center valign=middle><p>操作完成</p><a href='#' onclick='javascript:top.window.close();' onmouseover="&chr(34)&"window.status='关闭';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">关闭</a> <a href='onlinelist.asp?chatroomsn="&chatroomsn&"' onmouseover="&chr(34)&"window.status='返回';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">返回</a></div></body>"
%>

	