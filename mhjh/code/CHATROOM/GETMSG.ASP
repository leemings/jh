<%
Response.Expires=-1
chatroomsn=session("yx8_mhjh_userchatroomsn")
maxnosaytime=clng(Application("yx8_mhjh_nosaytime"))
onlinename=Application("yx8_mhjh_onlinename"&chatroomsn)
webicqname=Application("yx8_mhjh_webicqname")
songname=Application("yx8_mhjh_songname")
aberrantname=Application("yx8_mhjh_aberrantname")
adminname=Application("yx8_mhjh_admin")
username=Session("yx8_mhjh_username")
lastsavetime=Session("yx8_mhjh_userlastsavetime")
nowtime=now()
if instr(aberrantname,";"&username&";")<>0 then
aberrantlist=Application("yx8_mhjh_aberrantlist")
aberrantlistubd=ubound(aberrantlist)
dim newaberrantlist()
newaberrantname=";"
m=1
for i=1 to aberrantlistubd step 4
tt=datediff("s",aberrantlist(i+3),nowtime)
if aberrantlist(i)=username and tt<0 and aberrantlist(i+2)="死亡" then
session.Abandon
Response.Redirect "chaterror.asp?id=002&from="&aberrantlist(i+1)
elseif aberrantlist(i)=username and tt<0 and aberrantlist(i+2)="驱逐" then
session.Abandon
response.redirect "chaterror.asp?id=005&from="&aberrantlist(i+1)
elseif aberrantlist(i)=username and tt<0 and aberrantlist(i+2)="锁定" then
session.Abandon
response.redirect "chaterror.asp?id=008&from="&aberrantlist(i+1)
elseif aberrantlist(i)=username and tt<0 and aberrantlist(i+2)="逮捕" then
session.Abandon
response.redirect "chaterror.asp?id=006&from="&aberrantlist(i+1)
elseif aberrantlist(i)=username and tt<0 and aberrantlist(i+2)="入狱" then
session.Abandon
response.redirect "chaterror.asp?id=007&from="&aberrantlist(i+1)
elseif aberrantlist(i)=username and tt<0 and aberrantlist(i+2)="炸弹" then
session.Abandon
response.redirect "bomb.asp"
elseif aberrantlist(i)=username and tt<0 and aberrantlist(i+2)="中毒" then
			talkarr=Application("yx8_mhjh_talkarr")
			talkpoint=clng(Application("yx8_mhjh_talkpoint"))
			dim newtalkarr(600)
			j=1
			for k=11 to 600 step 10
				newtalkarr(j)=talkarr(k)
				newtalkarr(j+1)=talkarr(k+1)
				newtalkarr(j+2)=talkarr(k+2)
				newtalkarr(j+3)=talkarr(k+3)
				newtalkarr(j+4)=talkarr(k+4)
				newtalkarr(j+5)=talkarr(k+5)
				newtalkarr(j+6)=talkarr(k+6)
				newtalkarr(j+7)=talkarr(k+7)
				newtalkarr(j+8)=talkarr(k+8)
				newtalkarr(j+9)=talkarr(k+9)
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
			set conn=server.CreateObject("adodb.connection")
			conn.Open Application("yx8_mhjh_connstr")
			set rst=server.CreateObject("adodb.recordset")
			rst.Open "select 体力 from 用户 where 姓名='"&username&"'",conn
			redim preserve newaberrantlist(m),newaberrantlist(m+1),newaberrantlist(m+2),newaberrantlist(m+3)
			newaberrantlist(m)=aberrantlist(i)
			newaberrantlist(m+1)=aberrantlist(i+1)
			if rst("体力")<=abs(tt) then
				conn.Execute "update 用户 set 状态='死亡',最后登录时间=#"&dateadd("n",30,nowtime)&"# where 姓名='"&username&"'"
				newtalkarr(599)="<font color=FF0000>【中毒】</font>由于身中剧毒，直接导致##死亡<font class=timsty>"&time()&"</font>"
				newtalkarr(600)=chatroomsn
				newaberrantlist(m+2)="死亡"
				newaberrantlist(m+3)=dateadd("n",30,nowtime)
			else
				conn.Execute "update 用户 set 体力=体力+"&tt*10&" where 姓名='"&username&"'"
				newtalkarr(599)="<font color=FF0000>【中毒】</font>##身中剧毒未解，体力因而减少"&abs(tt*10)&"<font class=timsty>("&time()&")</font>"
				newtalkarr(600)=chatroomsn
				newaberrantlist(m+2)=aberrantlist(i+2)
				newaberrantlist(m+3)=aberrantlist(i+3)
			end if
			m=m+4
			newaberrantname=newaberrantname&aberrantlist(i)&";"
			rst.Close
			set rst=nothing	
			conn.Close
			set conn=nothing
			Application.Lock
			Application("yx8_mhjh_talkpoint")=talkpoint+1
			Application("yx8_mhjh_talkarr")=newtalkarr
			Application.UnLock
			erase newtalkarr
			erase talkarr
	elseif tt<0 then
redim preserve newaberrantlist(m),newaberrantlist(m+1),newaberrantlist(m+2),newaberrantlist(m+3)
newaberrantlist(m)=aberrantlist(i)
newaberrantlist(m+1)=aberrantlist(i+1)
newaberrantlist(m+2)=aberrantlist(i+2)
newaberrantlist(m+3)=aberrantlist(i+3)
m=m+4
newaberrantname=newaberrantname&aberrantlist(i)&";"
end if
next
Application.Lock
if m=1 then
dim index(0)
Application("yx8_mhjh_aberrantlist")=index
else
Application("yx8_mhjh_aberrantlist")=newaberrantlist
end if
Application("yx8_mhjh_aberrantname")=newaberrantname
Application.UnLock
erase newaberrantlist
erase aberrantlist
end if
if username="" or instr(onlinename,";"&username&";")=0 then response.redirect "chaterror.asp?id=000"
'if datediff("s",session("yx8_mhjh_userlasttalktime"),nowtime)>maxnosaytime then response.redirect "chaterror.asp?id=001"
dim showarr()
mytalkpoint=session("yx8_mhjh_usertalkpoint")
newtalkpoint=0
talkarr=Application("yx8_mhjh_talkarr")
j=1
for i=1 to 600 step 10
newtalkpoint=talkarr(i)
if newtalkpoint>mytalkpoint and cstr(talkarr(i+9))=cstr(chatroomsn) and (talkarr(i+2)="0" or username=adminname or talkarr(i+4)="大家" or talkarr(i+2)="1" and (talkarr(i+3)=username or talkarr(i+4)=username)) then
redim preserve showarr(j),showarr(j+1),showarr(j+2),showarr(j+3),showarr(j+4),showarr(j+5),showarr(j+6),showarr(j+7),showarr(j+8),showarr(j+9)
showarr(j)=talkarr(i)
showarr(j+1)=talkarr(i+1)
showarr(j+2)=talkarr(i+2)
showarr(j+3)=talkarr(i+3)
showarr(j+4)=talkarr(i+4)
showarr(j+5)=talkarr(i+5)
showarr(j+6)=talkarr(i+6)
showarr(j+7)=talkarr(i+7)
showarr(j+8)=talkarr(i+8)
showarr(j+9)=talkarr(i+9)
j=j+10
end if
next
Response.Write "<script language=javascript>setTimeout('this.location.reload();',10000);"
if j>1 then
for i=1 to ubound(showarr) step 10
Response.Write "parent.showmsg('"&showarr(i+1)&"','"&showarr(i+2)&"','"&showarr(i+3)&"','"&showarr(i+4)&"','"&showarr(i+5)&"','"&showarr(i+6)&"','"&showarr(i+7)&"','"&showarr(i+8)&"');"
next
end if
if instr(webicqname,";"&username&";")<>0 and songname<>username then
Response.Write "window.open('getwebicq.asp','webicq','top=100,left=100,height=200,width=400,status=no,scrollbars=yes');"
elseif instr(webicqname,";"&username&";")<>0 and songname=username then
Response.Write "window.open('g.asp','webicq','top=100,left=100,height=200,width=409,status=no,scrollbars=yes');"
elseif datediff("s",lastsavetime,nowtime)>300 then
Response.Write "parent.resfrm.location.replace('saveexp.asp');"
end if
Response.Write "</script>"
if newtalkpoint>mytalkpoint then session("yx8_mhjh_usertalkpoint")=newtalkpoint
%>
