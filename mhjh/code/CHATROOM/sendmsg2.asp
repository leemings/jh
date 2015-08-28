<%@ LANGUAGE=VBScript%>
<%
Response.Expires=-1
mycorp=session("yx8_mhjh_usercorp")
mygrade=Session("yx8_mhjh_usergrade")
chatroomsn=Session("yx8_mhjh_userchatroomsn")
maxnosaytime=Application("yx8_mhjh_nosaytime")
username=Session("yx8_mhjh_username")
title=Request.Form("title")
onlinename=Application("yx8_mhjh_onlinename"&chatroomsn)
aberrantname=Application("yx8_mhjh_aberrantname")
msg=server.HTMLEncode(trim(Request.Form("talkmsgtmp")))
namecolor=Request.Form("namecolor")
wordcolor=Request.Form("wordcolor")
sendto=Request.Form("sendto")
expression=Request.Form("expression")
xq=Request.Form("xq")
isprivacy=Request.Form("isprivacy")
act=0
if isprivacy="on" and sendto<>"大家" and sendto<>username then 
	isprivacy=1
else
	isprivacy=0
end if
if sendto="武术家" then 
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst1=server.CreateObject("adodb.recordset")
rst1.Open "select 体力,自动回复 from 用户 where 姓名='武术家'",conn
tili=rst1("体力")
huifu=rst1("自动回复")
rst1.close
set rst1=nothing
if tili>0 and huifu=true then 
talkarr1=Application("yx8_mhjh_talkarr")
talkpoint=clng(Application("yx8_mhjh_talkpoint"))
dim newtalkarr1(600)
j=1
for i=11 to 600 step 10
	newtalkarr1(j)=talkarr1(i)
	newtalkarr1(j+1)=talkarr1(i+1)
	newtalkarr1(j+2)=talkarr1(i+2)
	newtalkarr1(j+3)=talkarr1(i+3)
	newtalkarr1(j+4)=talkarr1(i+4)
	newtalkarr1(j+5)=talkarr1(i+5)
	newtalkarr1(j+6)=talkarr1(i+6)
	newtalkarr1(j+7)=talkarr1(i+7)
	newtalkarr1(j+8)=talkarr1(i+8)
	newtalkarr1(j+9)=talkarr1(i+9)
	j=j+10
next
newtalkarr1(591)=talkpoint+1
newtalkarr1(592)=1
newtalkarr1(593)=0
newtalkarr1(594)=username
newtalkarr1(595)=sendto
newtalkarr1(596)=""
newtalkarr1(597)=namecolor
newtalkarr1(598)="008000"
if instr(msg,"站长")<>0 then
newtalkarr1(599)="<font color=FF0000>〖武术家〗说道：“别轻易惹我，虽然我不愿意杀人，但你想死，我也没办法!你打我多少体力我不管，但你每次攻击我，我都要吸取你的<font color=FF0000>等级*2000</font>那么多的体力给我自己加上，有本事你就杀了我，我身上的东西随便你拿!但你从我身上抢到的东西的数量都是随机的，能有多少就看你的运气了!嘿嘿，来吧，用你学过武功出手吧，好处多多呢!”"
else 
newtalkarr1(599)="<font color=FF0000>〖武术家〗说道：</font>“别轻易惹我，虽然我不愿意杀人，但你想死，我也没办法!你打我多少体力我不管，但你每次攻击我，我都要吸取你的<font color=FF0000>等级*2000</font>那么多的体力给我自己加上，有本事你就杀了我，我身上的东西随便你拿!但你从我身上抢到的东西的数量都是随机的，能有多少就看你的运气了!嘿嘿，来吧，用你学过武功出手吧，好处多多呢!”"
end if 
newtalkarr1(600)=chatroomsn
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr1
Application.UnLock
end if
end if
act=0
if Application("yx8_mhjh_systemname"&chatroomsn)<>"官府" then
	if Application("yx8_mhjh_star"&chatroomsn)="" then Application("yx8_mhjh_star"&chatroomsn)="star"
	if application("questiontime"&chatroomsn)="" or (int(datediff("n",application("questiontime"&chatroomsn),now()))>=2 and Application("yx8_mhjh_star"&chatroomsn)="star") then 
dim conn,rst 
application.lock
application("questiontime"&chatroomsn)=now()
randomize timer
r=int(rnd*30)+1
		Application.Lock
		Application("yx8_mhjh_klt&r")=r
		Application.UnLock	
select case r
%><!--#include file="npc.asp"--><%
if username="" or instr(onlinename,";"&username&";")=0 then Response.Redirect "chaterror.asp?id=000"
if len(msg)>120 then msg=Left(msg,120)
if msg="" or msg="//" or msg="/#" then Response.End
if InStr(msg,"")<>0 or InStr(msg,"&#63736")<>0 then Response.End
nowtime=now()
nosaytime=datediff("s",session("yx8_mhjh_userlasttalktime"),nowtime)
if nosaytime<2 then
	Response.Write "<script language=JavaScript>{alert('有话慢慢说，别噎着了！');}</script>"
	Response.End
end if
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst1=server.CreateObject("adodb.recordset")
rst1.Open "select * from 门派 where 加入条件<>'False' and 门派='"&sendto&"'",conn
if rst1.EOF or rst1.BOF then
	Response.Write "<script language=JavaScript>{alert('这个门派你现在不能骚扰！');}</script>"
	Response.End
else
corp=rst1("门派")
rst1.close
set rst1=nothing
end if
if not (sendto="大家" or sendto="武术家" or sendto="魔法师"  or sendto=""&corp&"" or sendto="爱情鸟" or instr(onlinename,";"&sendto&";")<>0) then
	Response.Write "<script Language=Javascript>alert('"&sendto&"不在聊天室中，不能对其发言。');parent.resfrm.location.href='onlinelist.asp';</script>"
	Response.end
end if
if instr(msg,"C")<>0 or instr(msg,"你吗")<>0 or instr(msg,"犊子")<>0  or instr(msg,"狗")<>0  or instr(msg,"逼")<>0 or instr(msg,"B")<>0 or instr(msg,"奶")<>0 or instr(msg,"爷")<>0 or instr(msg,"爸")<>0 or instr(msg,"儿")<>0 or instr(msg,"娘")<>0 or instr(msg,"操")<>0 or instr(msg,"靠")<>0 or instr(msg,"妈")<>0 or instr(msg,"孙")<>0 or instr(msg,"MA")<>0 then
    Response.Write "<script language=JavaScript>{alert('对不起，禁止说带有骂人字眼的话。这次警告,下次抓!');}</script>"
    Response.End
end if
if instr(msg,"几吧")<>0 or instr(msg,"乌龟")<>0 or instr(msg,"日你")<>0  or instr(msg,"强奸")<>0  or instr(msg,"我日")<>0 or instr(msg,"叼")<>0 or instr(msg,"你奶")<>0 or instr(msg,"阴道")<>0 or instr(msg,"王八")<>0 or instr(msg,"牛比")<>0 or instr(msg,"你妈")<>0 or instr(msg,"操你")<>0 or instr(msg,"干你")<>0 or instr(msg,"你马")<>0 or instr(msg,"杂种")<>0 or instr(msg,"鸡吧")<>0 or instr(msg,"mm")<>0 then
    Response.Write "<script language=JavaScript>{alert('对不起，禁止脏话!这次让你退出聊天室,下次你就重新启动机器吧!!');parent.resfrm.location.href='../Exit.asp';}</script>"
    Response.End
end if
if title="on" then
  if mygrade<20 then
    Response.Write "<script language=JavaScript>{alert('发标题需要20级以上');}</script>"
    Response.write "<script language=javascript>parent.talkfrm.talkform.title.checked=false;</script>"
    Response.end
else
    Application.lock
    Application("yx8_mhjh_advertisemen")="<font color=red>"&msg&"</font><font color=red>（"&username&"-"&time&"）</font>"
    Application.unlock
    Response.write "<script language=javascript>parent.talkfrm.talkform.title.checked=false;</script>"
    Response.write "<script language=javascript>parent.titlefrm.location.replace('title.asp');</script>"
    Response.end
  end if
end if
if Instr(Application("yx8_mhjh_zlname"),";"&username&";")<>0 and Left(msg,4)<>"//回来" then
Response.Write "<script Language=Javascript>alert('你选择了离开座位，请先执行“//回来”，才可以再说话！');</script>" 
Response.end 
end if
if Instr(Application("yx8_mhjh_zlname"),";"&sendto&";")<>0 and Left(msg,4)<>"//回来" then
Response.Write "<script Language=Javascript>alert('对方暂时离开了座位,请稍后再和他联系');</script>" 
Response.end 
end if
session("yx8_mhjh_userlasttalktime")=nowtime
if instr(aberrantname,username)<>0 then
	abl=Application("yx8_mhjh_aberrantlist")
	ablubd=ubound(abl)
	for i=1 to ablubd step 4
		tt=datediff("s",abl(i+3),nowtime)
		if abl(i)=username and tt<0 and abl(i+2)="麻痹" then
			Response.Write "<script language=javascript>alert('麻痹中，动弹不得！');</script>"
			Response.End
		elseif abl(i)=username and tt<0 and abl(i+2)="点穴" and Left(msg,4)<>"//解穴" then
		    t=abs(tt)
			Response.Write "<script language=javascript>alert('你被点穴了,还敢动？等"&t&"秒');</script>"
			Response.End
		elseif abl(i)=username and tt<0 and abl(i+2)="中毒" then
			Response.Write "<script language=javascript>alert('中毒中，动弹不得！');</script>"
			Response.End
		elseif abl(i)=username and tt<0 and abl(i+2)="封咒"and left(msg,2)="//" then
			Response.Write "<script language=javascript>alert('封咒状态中，无法使用任何命令！');</script>"
			Response.End
		elseif abl(i)=username and tt<0 and abl(i+2)="疯狂" then
			olt=Application("yx8_mhjh_onlinelist")
			oltubd=ubound(olt)
			randomize()
			rnds=clng((oltubd\6-1)*rnd())
			sendto=olt(rnds*6+1)
		elseif abl(i)=username and tt<0 and instr(";死亡;逮捕;入狱;驱逐;锁定;",";"&abl(i+2)&";") then
		    session.Abandon
			Response.Redirect "chaterror.asp?id=000"
		end if	
	next 
end if
if left(msg,2)="//" then
%><!--#include file="ss.asp"--><%
elseif left(msg,2)="/#" then
	msg="##<font color="&wordcolor&">"&mid(msg,3)&"</font>"
end if
msg=replace(msg,"\","\\")
msg=replace(msg,"/","\/")
msg=replace(msg,"&lt"," ")
msg=replace(msg,"&gt"," ")
msg=replace(msg,"/x3c"," ")
msg=replace(msg,"\x3c"," ")
msg=replace(msg,"\x3e"," ")
msg=replace(msg,"/x3e"," ")
msg=replace(msg,"/x7d"," ")
msg=replace(msg,"\x7d"," ")
msg=Replace(msg,"\074"," ")
msg=Replace(msg,"\74"," ")
msg=Replace(msg,"\75"," ")
msg=Replace(msg,"\76"," ")
msg=Replace(msg,"\076"," ")
msg=Replace(LCase(msg),LCase("java"),"f2uck")
msg=Replace(LCase(msg),LCase("con"),"f2uck")
msg=replace(msg,chr(34),"\"&chr(34))
msg=replace(msg,chr(39),"\"&chr(39))
talkarr=Application("yx8_mhjh_talkarr")
talkpoint=clng(Application("yx8_mhjh_talkpoint"))
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
newtalkarr(592)=act
newtalkarr(593)=isprivacy
newtalkarr(594)=username
newtalkarr(595)=sendto
newtalkarr(596)=expression
newtalkarr(597)=namecolor
newtalkarr(598)=wordcolor
if xq="" then
newtalkarr(599)=msg&"<font class=timsty><\/font>"
else
newtalkarr(599)=msg&"<font class=timsty><\/font>"&xq&""
end if
newtalkarr(600)=chatroomsn
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
dim showarr()
mytalkpoint=session("yx8_mhjh_usertalkpoint")
newtalkpoint=0
talkarr=Application("yx8_mhjh_talkarr")
j=1
for i=1 to 600 step 10
	newtalkpoint=talkarr(i)
	if newtalkpoint>mytalkpoint and cstr(talkarr(i+9))=cstr(chatroomsn) and (talkarr(i+2)="0" or username=Application("yx8_mhjh_admin") or talkarr(i+4)="大家" or (talkarr(i+2)="1" and (talkarr(i+3)=username or talkarr(i+4)=username))) then
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
if j>1 then
Response.Write "<script language=javascript>"
	for i=1 to ubound(showarr) step 10
		Response.Write "parent.showmsg('"&showarr(i+1)&"','"&showarr(i+2)&"','"&showarr(i+3)&"','"&showarr(i+4)&"','"&showarr(i+5)&"','"&showarr(i+6)&"','"&showarr(i+7)&"','"&showarr(i+8)&"');"
	next
Response.Write "</script>"
end if
erase talkarr
erase showarr
if newtalkpoint>mytalkpoint then session("yx8_mhjh_usertalkpoint")=newtalkpoint
%>



