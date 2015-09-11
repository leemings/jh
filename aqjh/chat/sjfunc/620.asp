<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="func.asp"-->
<!--#include file="sjfunc.asp"-->
<%'暂离
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if aqjh_jhdj<jhdj_zl then
	Response.Write "<script language=JavaScript>{alert('提示：暂离需["&jhdj_zl&"]级以上！');}</script>"
	Response.End
end if
if chatinfo(11)=0 then
	Response.Write "<script language=JavaScript>{alert('提示：此房间禁上泡点，不允许暂离！');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
if towho<>Application("aqjh_automanname") and towho<>"大家" and Instr(Application("aqjh_useronlinename"&nowinroom)," "&towho&" ")=0 then
	Response.Write "<script Language=Javascript>alert('“" & towho & "”不在聊天室中，不能对其发言。');parent.f2.document.af.towho.value='大家';parent.f2.document.af.towho.text='大家';parent.m.location.reload();</script>"
	Response.end
end if
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【暂离】<font color=" & saycolor & ">"+zanli(mid(says,i+1))+"</font></font>"
towho="大家"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'暂离
function zanli(fn1)
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0 then
	Response.Write "<script language=JavaScript>{alert('您目前已经在暂离状态，请不要重复使用！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,内力,操作时间,会员等级,身份,门派,名单头像,性别,好友名单,配偶,国家,职位,通缉,转生,进化 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<3 and aqjh_grade<10 then
	s=3-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：您要进行暂离开功能请等["&s&"]分再操作！');}</script>"
	Response.End
end if
if rs("内力")<1000 then
	Response.Write "<script language=JavaScript>{alert('提示：需要内力1000、好好练几天吧！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用户 set 内力=内力-1000,操作时间=now() where 姓名='" & aqjh_name &"'"
aqjh_id=rs("id")
hydj=rs("会员等级")
jhsf=rs("身份")
if Instr(Application("aqjh_guibin"),"|" & aqjh_name & "|")<>0 then
 jhsf="贵宾"
end if
if Instr(Application("aqjh_admin_send"),"|" & aqjh_name & "|")<>0 then 
 jhsf="财神"
end if
if rs("配偶")=Application("aqjh_user") and rs("性别")="女" then
 jhsf="站长夫人"
end if
if Application("aqjh_mengzhu")=aqjh_name then 
 jhsf="武林盟主"
end if
if rs("通缉")=True then
	jhmp="通缉犯"
else
	jhmp=rs("门派")
end if
jhtx=rs("名单头像")
sex=rs("性别")
myzs=rs("转生")
mypeiou=rs("配偶")
guojia=rs("国家")
zhiwei=rs("职位")
jhjh1=rs("进化")
rs.close

if jhjh1<>"" then
jhjhinfo=split(jhjh1,"|||")
jhjhtitle=jhjhinfo(0)
jhjhdj=jhjhinfo(1)
if jhjhdj=0 then
jhjhimg=""
jhjhtitle="初学忍者"
elseif jhjhdj=1 then
jhjhimg="<img src=time_star.gif border=0>"
jhjhtitle="原始人"
elseif jhjhdj=2 then
jhjhimg="<img src=time_star.gif border=0><img src=time_star.gif border=0>"
jhjhtitle="岩洞人"
elseif jhjhdj=3 then
jhjhimg="<img src=time_star.gif border=0><img src=time_star.gif border=0><img src=time_star.gif border=0>"
jhjhtitle="稻草人"
elseif jhjhdj=4 then
jhjhimg="<img src=time_yueliang.gif border=0>"
jhjhtitle="木头人"
elseif jhjhdj=5 then
jhjhimg="<img src=time_yueliang.gif border=0><img src=time_star.gif border=0>"
jhjhtitle="小铁人"
elseif jhjhdj=6 then
jhjhimg="<img src=time_yueliang.gif border=0><img src=time_star.gif border=0><img src=time_star.gif border=0>"
jhjhtitle="白钢人"
elseif jhjhdj=7 then
jhjhimg="<img src=time_yueliang.gif border=0><img src=time_star.gif><img src=time_star.gif border=0><img src=time_star.gif border=0>"
jhjhtitle="现代人"
elseif jhjhdj=8 then
jhjhimg="<img src=time_yueliang.gif border=0><img src=time_yueliang.gif border=0>"
jhjhtitle="未来人"
elseif jhjhdj=9 then
jhjhimg="<img src=time_yueliang.gif border=0><img src=time_yueliang.gif border=0><img src=time_yueliang.gif border=0>"
jhjhtitle="太空人"
elseif jhjhdj=10 then
jhjhimg="<img src=jhjh/images/time_sun.gif border=0>"
jhjhtitle="宇宙人"
elseif jhjhdj=11 then
jhjhimg="<img src=time_sun.gif border=0><img src=time_sun.gif border=0>"
jhjhtitle="幻影人"
elseif jhjhdj=12 then
jhjhimg="<img src=time_sun.gif border=0><img src=time_sun.gif border=0><img src=time_sun.gif border=0>"
jhjhtitle="自由人"
elseif jhjhdj=13 then
jhjhimg="<img src=time_sun.gif border=0><img src=time_sun.gif border=0><img src=time_sun.gif border=0><img src=time_sun.gif border=0>"
jhjhtitle="全能人"
end if
else
else
jhjhimg=""
jhjhtitle="初学忍者"
end if

set rs=nothing
conn.close
set conn=nothing
Application.Lock
onlinelist=Application("aqjh_onlinelist"&nowinroom)
onlinenum=UBound(onlinelist)
for i=1 to onlinenum step 1
	onlinexx=split(onlinelist(i),"|")
	if onlinexx(0)=aqjh_name then
		aqjh_zm=aqjh_name&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&aqjh_jhdj&"|"&aqjh_id&"|"&hydj&"|1"&"|"&onlinexx(9)&"|"&mypeiou&"|"&myzs&"|"&jhjhtitle&"|"&jhjhimg&"|"&guojia&"|"&zhiwei
		onlinelist(i)=aqjh_zm
		exit for
	end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
application("aqjh_zanli")=application("aqjh_zanli")&"!"&aqjh_name&"!"&"|"&fn1&"|"&";"
Application.UnLock
zanli=aqjh_name&fn1
Response.Write "<script>parent.m.location.reload();parent.f2.document.af.addvalues.checked=true;<"&"/script>"
end function
%>