<!--#include file="czdj.asp"-->
<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
function bdsays(says)
	says=replace(says,"\","")
	says=replace(says,"&#","")
	says=replace(says,"..","")
	says=Replace(says,"<"," ")
	says=Replace(says,">"," ")
	says=Replace(says,"\x3c"," ")
	says=Replace(says,"\x3e"," ")
	says=Replace(says,"\074"," ")
	says=Replace(says,"\74"," ")
	says=Replace(says,"\75"," ")
	says=Replace(says,"\76"," ")
	says=Replace(says,"&lt"," ")
	says=Replace(says,"&gt"," ")
	says=Replace(says,"\076"," ")
	says=Replace(LCase(says),LCase("con"),"f2uck")
	says=Replace(LCase(says),LCase("or"),"ｏｒ")
	says=Replace(LCase(says),LCase("java"),"f2uck")
	says=Replace(LCase(says),LCase("www"),"f2uck")
	says=Replace(LCase(says),LCase("com"),"f2uck")
	says=Replace(LCase(says),LCase("net"),"f2uck")
	says=Replace(says,"Ｗ","f2uck")
	says=Replace(says,"ｗ","f2uck")
	says=Replace(says,"asp","f2uck")
	says=Replace(says,"jh","f2uck")
	says=Replace(says,"去了可以当","f2uck")
	says=Replace(says,"去了可以作","f2uck")
	says=Replace(says,"有一个好江湖","f2uck")
	says=Replace(says,"去我的江湖","f2uck")
	says=Replace(says,"想玩江湖","f2uck")
	says=Replace(says,"来我江湖","f2uck")
	says=Replace(LCase(says),LCase("http"),"f2uck")
	maren=instr(says,"f2uck")
	if maren<>0 then
		Response.Write "<script Language=Javascript>top.location.href='../autokick.asp';</script>"
		Response.end
	end if
bdsays=says
end function
'点穴，暂离
sub dianzan(towho)
'哑穴
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
	Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');}</script>" 
	Response.End 
end if 
if towho<>Application("sjjh_automanname") and towho<>"大家" and Instr(Application("sjjh_useronlinename"&nowinroom)," "&towho&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：“" & towho & "”不在聊天室中，不能对其发言。');parent.f2.document.af.towho.value='大家';parent.f2.document.af.towho.text='大家';parent.m.location.reload();</script>"
	Response.end
end if
Application.Lock
if Instr(LCase(application("sjjh_dianxuename")),LCase(sjjh_name))>0 then
dianxue=split(application("sjjh_dianxuename"),";")
for x=0 to UBound(dianxue)-1
	dxname=split(dianxue(x),"|")
if dxname(0)=sjjh_name then
	if DateDiff("s",dxname(1),now())<60 then
		Response.Write "<script Language=Javascript>alert('你已被哑穴，不能做任何事啊！还剩" & 60-DateDiff("s",dxname(1),now()) & "秒自动解开!');</script>"
		erase dianxue
		erase dxname
		response.end
	else
		dxcz=dianxue(x)&";"
		Application.Lock()
		application("sjjh_dianxuename")=replace(application("sjjh_dianxuename"),dxcz,"")
		Application.UnLock()
		Response.Write "<script Language=Javascript>alert('时间到了，你的哑穴自动解开了!');</script>"
		erase dianxue
		erase dxname
		response.end
	end if
end if
erase dxname
next
erase dianxue
end if
Application.UnLock
 '屏蔽
if Instr(LCase(application("qzld_pingbi")),LCase(";"&sjjh_name&"|"&towho&";"))>0 and left(says,4)<>"/屏蔽$" and sjjh_grade<7  then
 	Response.Write "<script Language=javascript>alert('您已经把人家屏蔽了，不能对他讲话!！');</script>"
  	response.end
end if
if Instr(LCase(application("qzld_pingbi")),LCase(";"&towho&"|"&sjjh_name&";"))>0 and sjjh_grade<7 then
 	Response.Write "<script Language=javascript>alert('对方已经把你屏蔽，你不能对他讲话！');</script>"
  	response.end
end if
'暂离开
if Instr(LCase(application("sjjh_zanli")),LCase("!"&sjjh_name&"!"))>0 and left(says,4)<>"/回来$" and sjjh_grade<10 then
	Response.Write "<script Language=Javascript>alert('您正在“暂离”状态中，请使用“我回来啦”功能解除！');</script>"
	response.end
end if
'判断对方操作
if Instr(LCase(application("sjjh_zanli")),LCase("!"&towho&"!"))>0 and left(says,4)<>"/回来$" then
	sjjh_zanli=split(application("sjjh_zanli"),";")
	for x=0 to UBound(sjjh_zanli)
		dxname=split(sjjh_zanli(x),"|")
	if dxname(0)="!"&towho&"!" then
		Response.Write "<script Language=Javascript>alert('"&towho&"说："&dxname(1)&"');</script>"
		erase dxname
		erase sjjh_zanli
		response.end
	end if
	erase dxname
	next
	erase sjjh_zanli
end if
end sub
'发言模块
sub chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
zj="<a href=javascript:parent.sw(\'[" & sjjh_name & "]\'); target=f2><font color="&addwordcolor&">"& sjjh_name & "</font></a>"
br="<a href=javascript:parent.sw(\'[" & towho & "]\'); target=f2><font color="&addwordcolor&">"& towho & "</font></a>"
says=Replace(says,"##",zj)
says=Replace(says,"%%",br)
says=Replace(says,"$$redb","<font color=red><b>")
says=Replace(says,"$$blueb","<font color=blue><b>")
says=Replace(says,"$$yellowb","<font color=yellow><b>")
says=Replace(says,"$$greenb","<font color=green><b>")
says=Replace(says,"$$b","</b></font>")
says=Replace(says,"$$red","<font color=red>")
says=Replace(says,"$$blue","<font color=blue>")
says=Replace(says,"$$yellow","<font color=yellow>")
says=Replace(says,"$$green","<font color=green>")
says=Replace(says,"$$","</font>")
saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway & "," & nowinroom & ");<"&"/script>"
AddMsg SayStr
For i = 1 to Application("SayCount")-Session("SayCount")
	Response.Write Application("SayStr"&YuShu((Session("SayCount")+i)))
Next
session("SayCount")=Application("SayCount")
Response.Write "<script Language=JavaScript>"
Response.Write "parent.f1.scrollWindow();"
Response.Write "</script>"
end sub
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
	Application.Lock()
	Application("SayCount")=Application("SayCount")+1
	i="SayStr"&YuShu(Application("SayCount"))
	Application(i)=Str
	Application.UnLock()
End Sub
sub myclose(fs)
select case fs
case 0
	 if IsObject(rs) then
		set rs=nothing
	 end if
	 if IsObject(conn) then
	 	set conn=nothing
	 end if
case 1
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
case 2
	rs.close
	set rs=nothing
case 3
	conn.close
	set conn=nothing
end select
end sub
sub mess(mymess,fs)
Response.Write "<script language=JavaScript>{alert('"&mymess&"');}</script>"
Response.End
end sub

function ErrSays(from1,strContent)
ErrSays="[<a href=javascript:parent.sw('["&from1&"]'); target=f2> "&from1 & " </a>]" & strContent
end function

Function ErrALT2(strContent)
ErrALT2 = "<script Language=Javascript>alert('" & strContent & "');</script>"
End Function

Sub ErrALT(strContent)
Response.Write "<script Language=Javascript>alert('" & strContent & "');window.close();</script>"
Response.end
End Sub

%>