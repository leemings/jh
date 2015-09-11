<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'小孩护主
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
'#####房间处理#####
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=xhhz(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'小孩护主
function xhhz(fn1,to1)
fn1=trim(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select boysex,boy,配偶,hw FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,3,3
mypic="<img src=/chat/boy/"&rs("boysex")&" width=30 height=30>"
mypo=rs("配偶")
myhw=rs("hw")
cwnl="体力"
rs.close
rs.open "select "& cwnl &",门派,等级,保护,grade,宝物,死亡时间 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("门派")="出家" and aqjh_grade<10 then
rs.close
set rs=nothing 
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：他是出家人不可以操作！');}</script>"
Response.End
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<5 and rs("宝物")="无" and aqjh_grade<10 then
rs.close
set rs=nothing 
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：他刚被别人杀死，先放过他吧！');}</script>"
Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：请不要对官府进行攻击！！');}</script>"
Response.End
end if
if rs("等级")<=18 and rs("宝物")="无" and aqjh_grade<10 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：请不要对初入江湖新手操作！');}</script>"
Response.End
end if
if rs("保护")=True and aqjh_grade<10 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：对方正在练功保护请不要偷袭!');}</script>"
Response.End
end if
rs.close
rs.open "select 死亡时间,门派,grade,保护,杀人数 from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
sj=DateDiff("n",rs("死亡时间"),now())
if sj<15 and aqjh_grade<10 then
rs.close
set rs=nothing 
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：你刚刚被别人杀死，还是先练一会吧！');}</script>"
Response.End
end if
if rs("门派")="出家" and aqjh_grade<10 then
rs.close
set rs=nothing 
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：你是出家人不可以操作！');}</script>"
Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：你是官府中人请不要发招攻击！！！');}</script>"
Response.End
end if
if rs("保护")=True and aqjh_grade<10 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：你正在练功保护请不要发招!');}</script>"
Response.End
end if
if rs("杀人数")>=int(Application("aqjh_killman")) and aqjh_grade<10 then
conn.execute "update 用户 set 保护=false where 姓名='" & aqjh_name &"'"
Response.Write "<script language=JavaScript>{alert('提示：你作恶太多，超过江湖杀人上限"& Application("aqjh_killman") &"不能再发招了！');}</script>"
rs.close
set rs=nothing 
conn.close
set conn=nothing
Response.End
end if
if isnull(myhw) or myhw="" then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：您还没有小孩，请去领或生一个吧！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
response.end
end if
zt=split(myhw,"|")
if UBound(zt)<>12 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：小孩数据出错！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
response.end
end if
if zt(0)<>fn1 then
 Response.Write "<script Language=Javascript>alert('提示：小孩名字不对！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
response.end
end if
cssj=clng(DateDiff("d",zt(6),now()))
gs=clng(zt(4))
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
Response.Write "<script Language=Javascript>alert('提示：小孩心情不好，不想与你一起打架！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
response.end
end if
if clng(zt(2))<100 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：小孩体力虚弱，不能打架了！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
response.end
end if
rs.close
'姓名0|性别1|体力2|内力3|攻击4|防御5|生日6|状态7|喂养次数8|最后喂养9|感受10|头像11|最后查看时间12
cwsr=clng(DateDiff("d",zt(6),now()))
if cwsr<5 then
xhhz="<font color=green>【小孩护主】<font color=" & saycolor & ">##想使用小孩护主%%，可以小孩还是太小了，你忍心吗？(说明：同归于尽攻击计算,小孩出生天数除以5取整数，再去除攻击特性来攻击对方，攻击值给主人增加!)</font>"
exit function
end if
xhhzz=int(cwsr/5)
tgyjz1=int(tgyjz/xhhzz)
tgyjz2=tgyjz-tgyjz1
'conn.Execute ("update 用户 set "&cwnl&"="&tgyjz1&" where 姓名='" & to1 &"'")
'conn.Execute ("update 用户 set "&cwnl&"="&cwnl&"+"&tgyjz2&",cw='' where 姓名='" & aqjh_name &"'")
'conn.Execute ("update 用户 set 体力="&tgyjz1&" where 姓名='" & to1 &"'")
xhhz="<font color=green>【小孩护主】<font color=" & saycolor & ">##<bgsound src=wav/tgyj.wav loop=1>使用小孩[<b>"&zt(0)&"</b>]攻击%%，使%%体力下降到:<font color=red>-"&tgyjz1&"</font>点,小孩护主攻击(伤害:"& xhhzz &"),全身的血肉都化做一道道利箭向%%射去。%%悴不及防，被吸走[<b>"&cwnl&"</b>]:-"&tgyjz2&"点，小孩护主舍身取义，临死前的最后呐喊久久回荡在聊天室里…………</font>"
set rs=nothing
conn.close
set conn=nothing
end function
%>
