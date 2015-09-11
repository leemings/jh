<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'宠物搏斗
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<!--#include file="pkif.asp"-->
<%
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
says=Replace(says,"&","&")
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=cwbd(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'宠物攻击
function cwbd(fn1,to1)
fn1=trim(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
fn1=trim(fn1)
rs.open "select k,i FROM b WHERE a='" & fn1 &"'",conn
If Rs.Bof OR Rs.Eof then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('提示：你的宠物在系统数据库中并不存在\n请删除此物品或找管理员！！');}</script>"
 Response.End
end if
cwnl=rs("k")
mypic="<img src=../hcjs/jhjs/images/"&rs("i")&">"
rs.close
rs.open "select 死亡时间,门派,grade,保护,杀人数,cw from 用户 where  姓名='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("死亡时间"),now())
if sj<15 and aqjh_grade<10 then
 rs.close
 set rs=nothing 
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('提示：你刚刚被别人杀死，还是先练一会吧！');}</script>"
 Response.End
end if
if rs("门派")="出家" and aqjh_grade<10 then
 rs.close
 set rs=nothing 
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('提示：你是出家人不可以操作！');}</script>"
 Response.End
end if

if rs("grade")>=6 and rs("grade")<10 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('提示：你是官府中人请不要发招攻击！！！');}</script>"
 Response.End
end if
if rs("保护")=True and aqjh_grade<10 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('提示：你正在练功保护请不要发招!');}</script>"
 Response.End
end if
if rs("杀人数")>=int(Application("aqjh_killman"))  and aqjh_grade<10 then
 conn.execute "update 用户 set 保护=false where 姓名='" & aqjh_name &"'"
 Response.Write "<script language=javascript>{alert('提示：你作恶太多，超过江湖杀人上限"& Application("aqjh_killman") &"不能再发招了！');}</script>"
 rs.close
 set rs=nothing 
 conn.close
 set conn=nothing
 Response.End
end if
if isnull(rs("cw")) or rs("cw")="" then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('提示：没有宠物，请去购买！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
 response.end
end if
zt=split(rs("cw"),"|")
if UBound(zt)<>7 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('提示：宠物已经被打死，请重新购买！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
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
 Response.Write "<script Language=javascript>alert('提示：宠物心情不好，不想与你一起打架！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
 response.end
end if

rs.close
'名字0|类别1|生日2|体力3|攻击4|心情5|照顾数次6|照顾时间7
temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-int(clng(zt(4))/100))&"|"&(clng(zt(4))-int(clng(zt(4))/200))&"|"&zt(5)&"|"&zt(6)&"|"&now()
rs.open "select 门派,等级,保护,grade,宝物,死亡时间,cw from 用户 where 姓名='" & to1 &"'",conn
if isnull(rs("cw")) or rs("cw")="" then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('提示：对方还没有宠物，无法攻击！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
 response.end
end if
zt=split(rs("cw"),"|")
if UBound(zt)<>7 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('提示：对方宠物已死，无法攻击！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
 response.end
end if
if rs("门派")="出家" and aqjh_grade<10 then
 rs.close
 set rs=nothing 
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('提示：他是出家人不可以操作！');}</script>"
 Response.End
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<5 and rs("宝物")="无" and aqjh_grade<10 then
 rs.close
 set rs=nothing 
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('提示：他刚被别人杀死，先放过他吧！');}</script>"
 Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('提示：请不要对官府进行攻击！！');}</script>"
 Response.End
end if
if rs("等级")<=18 and rs("宝物")="无" and aqjh_grade<10 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('提示：请不要对初入江湖新手操作！');}</script>"
 Response.End
end if
if rs("保护")=True and aqjh_grade<10 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('提示：对方正在练功保护请不要偷袭!');}</script>"
 Response.End
end if
rs.close
temp1=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-int(clng(zt(4))/20))&"|"&(clng(zt(4))-int(clng(zt(4))/40))&"|"&zt(5)&"|"&zt(6)&"|"&now()
conn.Execute ("update 用户 set cw='"&temp1&"' where 姓名='" & to1 &"'")
conn.Execute ("update 用户 set cw='"&temp&"' where 姓名='" & aqjh_name &"'")
rs.open "select cw from 用户 where 姓名='" & to1 &"'",conn
zt=split(rs("cw"),"|")
if clng(zt(3))<-100 then
 conn.Execute ("update 用户 set cw='|' where 姓名='" & to1 &"'")
 conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','" & to1 & "的宠物让"&aqjh_name&"宠物打死','人命')"
 cwbd="<font color=green>【宠物搏斗】<font color=" & saycolor & ">##使用自己的宠物"&mypic&"攻击%%的宠物[<b>"&zt(0)&"</b>]，使[<b>"&zt(0)&"</b>]体力下降:<font color=red>-"&int(clng(zt(4))/20)&"</font>点,攻击下降:<font color=red>-"&int(clng(zt(4))/40)&"</font>点，%%的宠物终于体力不支倒地身亡……</font>"
else
 cwbd="<font color=green>【宠物搏斗】<font color=" & saycolor & ">##使用自己的宠物"&mypic&"攻击%%的宠物[<b>"&zt(0)&"</b>]，使[<b>"&zt(0)&"</b>]体力下降:<font color=red>-"&int(clng(zt(4))/20)&"</font>点，攻击下降:<font color=red>-"&int(clng(zt(4))/40)&"</font>点</font>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>