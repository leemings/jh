<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'夺宝宠物自暴攻击♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
call roompd("夺宝宠物自暴")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
towho=Trim(Request.Form("towho"))
if towho=sjjh_name or towho="大家" or towho=application("sjjh_automanname") then
	Response.Write "<script language=JavaScript>{alert('提示：不可以对自已或对机器人、大家使用！');}</script>"
	Response.End
end if
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
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
says=cwgj(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'夺宝宠物自暴攻击
function cwgj(fn1,to1)
fn1=trim(fn1)
toname=to1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select k,i FROM b WHERE a='" & fn1 &"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的宠物在系统数据库中并不存在\n请删除此物品或找管理员！');}</script>"
	Response.End
end if
cwnl=rs("k")
mypic="<img src=../hcjs/jhjs/images/"&rs("i")&">"
rs.close
rs.open "select id,"& cwnl &",门派,等级,保护,grade,宝物,死亡时间 from 用户 where 姓名='" & to1 &"'",conn,2,2
tgyjz=rs(cwnl)
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：请不要对官府进行攻击！');}</script>"
	Response.End
end if
tonameid=rs("id")
rs.close
rs.open "select id,等级,门派,grade,cw from 用户 where  姓名='"&sjjh_name&"'",conn,3,3
if rs("grade")>=6 and rs("grade")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是官府中人请不要发招攻击！');}</script>"
	Response.End
end if
if isnull(rs("cw")) or rs("cw")="" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您还没有宠物，请去购买！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
zt=split(rs("cw"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：宠物数据出错，请重新购买！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
cssj=clng(DateDiff("d",zt(2),now()))
gs=clng(zt(5))
if gs<(80+cssj*5) then
	czzt="叛逆"
elseif gs>((100+cssj*5+50)) then
	czzt="顺从"
else
	czzt="普通"
end if
if czzt="叛逆" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：宠物心情不好，不想与你一起打架！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if clng(zt(3))<100 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：宠物体力虚弱，不能打架了！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
'名字0|类别1|生日2|体力3|攻击4|心情5|照顾数次6|照顾时间7
cwsr=clng(DateDiff("d",zt(2),now()))
if cwsr<5 then
	cwgj="<font color=green>【夺宝宠物自暴】<font color=" & saycolor & ">##想使用宠物自暴攻击%%，可以宠物还在太，你忍心吗？(说明：同归于尽攻击计算,宠物出生天数除以5取整数，再去除攻击特性来攻击对方，攻击值给主人增加！)</font>"
	exit function
end if
cwgjz=int(cwsr/5)
tgyjz1=int(tgyjz/cwgjz)
tgyjz2=tgyjz-tgyjz1
conn.Execute ("update 用户 set "&cwnl&"="&tgyjz1&" where 姓名='" & to1 &"'")
conn.Execute ("update 用户 set "&cwnl&"="&cwnl&"+"&tgyjz2&",cw='' where 姓名='" & sjjh_name &"'")
conn.Execute ("update 用户 set 体力="&tgyjz1&" where 姓名='" & to1 &"'")
cwgj="<font color=green>【夺宝宠物自暴】<font color=" & saycolor & ">##使用宠物[<b>"&zt(0)&"</b>]"&mypic&"攻击%%，使%%体力下降到:<font color=red>-"&tgyjz1&"</font>点,宠物特技攻击(伤害:"& cwgjz &"),%%被吸走[<b>"&cwnl&"</b>]:-"&tgyjz2&"点，宠物护主舍身取义……</font>"
set rs=nothing
conn.close
set conn=nothing
end function
%>
