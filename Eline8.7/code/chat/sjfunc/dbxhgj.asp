<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'夺宝小孩攻击♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
'#####房间处理#####
call roompd("夺宝小孩攻击")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
if towho=sjjh_name or towho="大家" or towho=application("sjjh_automanname") then
	Response.Write "<script language=JavaScript>{alert('提示：不可以对自已或对机器人、大家使用！');}</script>"
	Response.End
end if
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=xhgj(towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'夺宝小孩攻击
function xhgj(to1)
fn1=trim(fn1)
toname=to1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")

rs.open "select id,等级,门派,grade,性别,配偶,hw from 用户 where 姓名='"&sjjh_name&"'",conn,3,3
if rs("保护")=true and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你正在练功保护中，不能攻击！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if rs("配偶")="" or rs("配偶")="无" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你还没结婚，那来的孩子！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if rs("等级")<30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：30级以上方可使用小孩攻击！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if sjjh_grade>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你是官府中人，不可以攻击！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if

mysex=rs("性别")
mypo=rs("配偶")
myhw=rs("hw")
mynameid=rs("id")
rs.close
if mypo=to1 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('不可以对自已的配偶进行此操作！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if mysex="女" then
	fyname=sjjh_name
else
	fyname=mypo
end if
if mysex="男" then
	rs.open "select hw from 用户 where 姓名='"&mypo&"'",conn,3,3
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：江湖中找不到你的配偶资料！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
		response.end
	end if
	myhw=rs("hw")
	rs.close
end if
if isnull(myhw) or myhw="" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你还没有小孩呢！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
zt=split(myhw,"|")
if UBound(zt)<>12 then
	erase zt
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：孩子数据出错！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if

xhname=zt(0) '小孩姓名
xhsex=zt(1) '性别
xhtl=clng(zt(2)) '体力
xhnl=clng(zt(3)) '内力
xhgj=clng(zt(4)) '攻击
xhfy=clng(zt(5)) '防御
xhday=zt(6) '生日
xhzt=zt(7) '状态
xhwy=clng(zt(8)) '喂养次数
xhwyrq=zt(9) '最后一次的喂养时间
xhgs=clng(zt(10)) '感受
xhtx=zt(11)
xhck=zt(12)
erase zt
if xhzt="死" or xhtl<=0 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的孩子已经死了，请先去复活！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
cssj=clng(DateDiff("d",xhday,now()))
if cssj<=30 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('小孩出生还不够30天，不能打架！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if xhgs<(80+cssj*5) then
	czzt="孤独"
elseif xhgs>((100+cssj*5+50)) then
	czzt="幸福"
else
	czzt="快乐"
end if
if czzt="孤独" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的小孩倍感孤独，不会和你一起打架的！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if xhtl<300 or xhnl<100 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的小孩体力虚弱，已不能再打架了！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.open "select id,门派,等级,grade,体力,内力,武功,攻击,防御 from 用户 where 姓名='" & to1 &"'",conn,3,3
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：请不要对官府进行攻击！！');}</script>"
	Response.End
end if
totl=rs("体力")
tonameid=rs("id")
rs.close
tlkiller=int((xhgj+xhfy)/2)
wgkiller=int((xhtl+xhnl)/8)
xhtl=int(xhtl*19/20)
xhnl=int(xhnl*19/20)
xhgj=xhgj-5
xhfy=xhfy-5
xhgs=xhgs-10
temp=xhname&"|"&xhsex&"|"&xhtl&"|"&xhnl&"|"&xhgj&"|"&xhfy&"|"&xhday&"|"&xhzt&"|"&xhwy&"|"&xhwyrq&"|"&xhgs&"|"&xhtx&"|"&xhck
conn.execute "update 用户 set 体力=体力-"&tlkiller&",武功=武功-"&wgkiller&" where 姓名='"&to1&"'"
conn.execute "update 用户 set hw='"&temp&"' where 姓名='"&fyname&"'"
conn.execute "update 用户 set 武功=武功+"&wgkiller&" where 姓名='"&sjjh_name&"'"
newtotl=totl-tlkiller
xhgj="##喊出自已的孩子<img src='../../ico/"& xhtx &"-2.gif'><font color=red>"&xhname&"</font>攻击%%，"
fjh=""
if newtotl<-100 then
	conn.execute "update 用户 set 状态='死' where 姓名='"&to1&"'"
	conn.execute "update 用户 set 杀人数=杀人数+1 where 姓名='"&sjjh_name&"'"
	call boot(to1,"发招，操作者："&sjjh_name&",["&xhname&"]小孩攻击打死")
	fjh="杀伤体力："&tlkiller&"点，吸取武功："&wgkiller&"点。%%两眼一翻，双脚一蹬，被黑白无常给收了去。"
	bbb=zxrsbd(sjjh_name,mynameid)
	newfjh=xhgj&fjh
	newfjh=replace(newfjh,"##",sjjh_name)
	newfjh=replace(newfjh,"%%",toname)
	newfjh=newfjh&"<BR><BR>"&bbb
	fjh=fjh&"<BR><BR>"&bbb
	call dbxx(newfjh)
	to1=toname
else
	fjh="杀伤体力："&tlkiller&"点，吸取武功："&wgkiller&"点。"
end if
xhgj=xhgj&fjh
set rs=nothing
conn.close
set conn=nothing
end function
%>