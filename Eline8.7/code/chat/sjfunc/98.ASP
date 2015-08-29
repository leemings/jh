<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'小孩攻击♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
'#####房间处理#####
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
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
'对暂离开、点哑穴判断
call dianzan(towho)
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

'宠物攻击
function cwgj(fn1,to1)
fn1=trim(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
fn1=trim(fn1)

rs.open "select boysex,boy,配偶 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,3,3
mypic="<img src=/chat/boy/"&rs("boysex")&" width=30 height=30>"
cwnl="内力"
zz=rs("配偶")

rs.close
rs.open "select 门派,等级,保护,grade,宝物,死亡时间 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("门派")="出家" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他是出家人不可以操作！');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<5 and rs("宝物")="无" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他刚被别人杀死，先放过他吧！');}</script>"
	Response.End
end if
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：请不要对官府进行攻击！！');}</script>"
	Response.End
end if
if rs("等级")<=12 and rs("宝物")="无" and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：请不要对初入江湖新手操作！');}</script>"
	Response.End
end if
if rs("保护")=True and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方正在练功保护请不要偷袭!');}</script>"
	Response.End
end if
rs.close
rs.open "select 门派,grade,保护,杀人数,boy from 用户 where  姓名='"&sjjh_name&"'",conn,2,2
if rs("门派")="出家" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是出家人不可以操作！');}</script>"
	Response.End
end if
if rs("grade")>=6 and rs("grade")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是官府中人请不要发招攻击！！！');}</script>"
	Response.End
end if
if rs("保护")=True and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你正在练功保护请不要发招!');}</script>"
	Response.End
end if
if rs("杀人数")>=int(Application("sjjh_killman"))  and sjjh_grade<10 then
	conn.execute "update 用户 set 保护=false where 姓名='" & sjjh_name &"'"
	Response.Write "<script language=JavaScript>{alert('提示：你作恶太多，超过江湖杀人上限"& Application("sjjh_killman") &"不能再发招了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("boy")="" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您还没有小孩，赶快生一个！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
zt=split(rs("boy"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：小孩数据出错！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if fn1<>zt(0) then
	Response.Write "<script language=JavaScript>{alert('提示：你怎么叫自己的孩子都叫错！');}</script>"
	Response.End
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
	Response.Write "<script Language=Javascript>alert('提示：小孩心情不好，不想与你一起打架！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if

if clng(zt(3))<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：小孩体力虚弱，不能打架了！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
'名字0|类别1|生日2|体力3|攻击4|心情5|照顾数次6|照顾时间7
temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-int(clng(zt(3))/10))&"|"&zt(4)&"|"&zt(5)&"|"&zt(6)&"|"&now()
conn.Execute ("update 用户 set "&cwnl&"="&cwnl&"-"&int(clng(zt(4))/10)&" where 姓名='" & to1 &"'")
conn.Execute ("update 用户 set "&cwnl&"="&cwnl&"+"&int(clng(zt(4))/10)&",boy='"&temp&"' where 姓名='" & sjjh_name &"'")
conn.Execute ("update 用户 set "&cwnl&"="&cwnl&"+"&int(clng(zt(4))/10)&",boy='"&temp&"' where 姓名='" & zz &"'")
conn.Execute ("update 用户 set 体力=体力-"&clng(zt(4))&" where 姓名='" & to1 &"'")
rs.open "select 体力 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("体力")<-100 then
	call boot(to1,"发招，操作者："&sjjh_name&",["&zt(0)&"]打死"&fn1)
	cwgj="<font color=green>【小孩攻击】<font color=" & saycolor & ">##叫自己小孩<b>"&zt(0)&"</b>]攻击%%，使%%体力下降:<font color=red>-"&clng(zt(4))&"</font>点,小孩特技攻击,%%被吸走[<b>"&cwnl&"</b>]:-"&int(clng(zt(4))/10)&"点,%%终于体力不支倒地，被["&zt(0)&"]打得七零八落………</font>"
else
	cwgj="<font color=green>【小孩攻击】<font color=" & saycolor & ">##叫自己小孩[<b>"&zt(0)&"</b>]攻击%%，使%%体力下降:<font color=red>-"&clng(zt(4))&"</font>点,小孩特技攻击,%%被吸走[<b>"&cwnl&"</b>]:-"&int(clng(zt(4))/10)&"点</font>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
