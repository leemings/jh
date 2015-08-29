<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="func.asp"-->
<!--#include file="sjfunc.asp"-->
<!--#include file="chatconfig.asp"-->
<%'暂离♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
fjname=chatinfo(0)
erase sjjh_roominfo
erase chatinfo
if Weekday(date())=7 and (Hour(time())>=20 and Hour(time())<21) and fjname="高手E线" then
	Response.Write "<script language=JavaScript>{alert('提示：夺宝大赛中不可以使用暂离功能！');}</script>"
	Response.End
end if
if sjjh_jhdj<jhdj_zl then
	Response.Write "<script language=JavaScript>{alert('提示：暂离功能需要["&jhdj_zl&"]级以上才可以使用！');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
if towho<>Application("sjjh_automanname") and towho<>"大家" and Instr(Application("sjjh_useronlinename"&nowinroom)," "&towho&" ")=0 then
	Response.Write "<script Language=Javascript>alert('“" & towho & "”不在聊天室中，不能对其发言。');parent.f2.document.af.towho.value='大家';parent.f2.document.af.towho.text='大家';parent.m.location.reload();</script>"
	Response.end
end if
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【暂时离开】<font color=" & saycolor & ">"+zanli(mid(says,i+1))+"</font></font>"
towho="大家"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'暂离
function zanli(fn1)
if Instr(LCase(application("sjjh_zanli")),LCase("!"&sjjh_name&"!"))>0 then
	Response.Write "<script language=JavaScript>{alert('您目前已经在暂离状态，请不要重复使用！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select id,内力,等级,操作时间,会员等级,身份,门派,名单头像,性别,好友名单,配偶,通缉 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
sj=DateDiff("n",rs("操作时间"),now())

if rs("等级")>12 and rs("内力")<1000 then
	Response.Write "<script language=JavaScript>{alert('提示：暂离需要内力1000、先练一下吧！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用户 set 内力=内力-1000,操作时间=now() where 姓名='" & sjjh_name &"'"
sjjh_id=rs("id")
hydj=rs("会员等级")
jhsf=rs("身份")
if rs("通缉")=True then
	jhmp="通缉犯"
else
	jhmp=rs("门派")
end if
jhtx=rs("名单头像")
sex=rs("性别")
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
onlinelist=Application("sjjh_onlinelist"&nowinroom)
onlinenum=UBound(onlinelist)
for i=1 to onlinenum step 1
	onlinexx=split(onlinelist(i),"|")
	if onlinexx(0)=sjjh_name then
		sjjh_zm=sjjh_name&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&sjjh_jhdj&"|"&sjjh_id&"|"&hydj&"|1"&"|"&onlinexx(9)
		onlinelist(i)=sjjh_zm
		exit for
	end if
next
Application("sjjh_onlinelist"&nowinroom)=onlinelist
application("sjjh_zanli")=application("sjjh_zanli")&"!"&sjjh_name&"!"&"|"&fn1&"|"&";"
Application.UnLock
zanli=sjjh_name&fn1
Response.Write ("<Script Language=JavaScript>parent.m.location.reload();</Script>")
end function
%>
<Script >
parent.crm='<%=Application("sjjh_chatroomname"&session("nowinroom"))%>';
parent.myroom=<%=session("nowinroom")%>
parent.f2.document.af.mdsx.checked=true;
parent.f2.document.af.sytemp.focus();
parent.m.location.reload();
parent.f2.location.href='../f22.asp?id=1';
</script>