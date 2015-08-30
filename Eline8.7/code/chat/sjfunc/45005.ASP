<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'寻找法器♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
call dianzan(towho)
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
says="<font color=red>【寻找法器】<font color=" & saycolor & ">"+xunfaqi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'寻找法器
function xunfaqi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 等级,法力,操作时间 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<(int(rnd*5)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*5)+1)-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
dj=rs("等级")

fla=rs("法力")
if rs("法力")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的法力不够无法施展呀，至少也得50000点啊！');window.close();}</script>"
	response.end
end if
if dj<50 then
Response.Write "<script language=JavaScript>{alert('此功能需要50级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

rs.close
conn.execute "update 用户 set 法力=法力-50000,操作时间=now()  where 姓名='"&sjjh_name&"'"
randomize 
r=int(rnd*9)+1
select case r
case 1
xunfaqi=sjjh_name & "你闯进" & to1 & "家里把<font color=red>法器铁笔银勾</font>从其手中抢走，真是世事不公啊."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"铁笔银勾",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close


case 2
xunfaqi=sjjh_name & "你在" & to1 & "家里的床底下偷到了<font color=red>天堂令</font>随后钻地洞跑了."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"天堂令",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close

	
case 3
xunfaqi=sjjh_name & "你在江湖里的一个破庙中发现了被别人丢弃的<font color=red>法器玄冥棒</font>."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"玄冥棒",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
	
case 4
xunfaqi=sjjh_name & "你真是幸运，难得一见的<font color=red>法器盗取令</font>都能被你找，真是有福之人啊."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"盗取令",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
	
case 5
xunfaqi=sjjh_name & "你发现" & to1 & "口袋里红光闪闪,顺手一摸原来是一颗<font color=red>红钻石</font>于是捡了块碎石头塞进" & to1 & "的口袋,把红钻石给盗走了.."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"红钻石",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close



case 6
xunfaqi=sjjh_name & "你发现" & to1 & "口袋里红光闪闪,顺手一摸原来是一颗<font color=red>绿钻石</font>于是捡了块碎石头塞进" & to1 & "的口袋,把绿钻石给盗走了.."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"绿钻石",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close

case 7
xunfaqi=sjjh_name & "你发现" & to1 & "口袋里红光闪闪,顺手一摸原来是一颗<font color=red>蓝钻石</font>于是捡了块碎石头塞进" & to1 & "的口袋,把蓝钻石给盗走了.."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"蓝钻石",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close

case 8
xunfaqi=sjjh_name & "你在江湖里的一棵千年古树的树枝上发现失传已久的<font color=red>九阳神功</font>" & sjjh_name & "这本九阳神功威力无比可让你成为武林第一高手"
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"九阳神功",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
	
case 9
xunfaqi=sjjh_name & "你发现" & to1 & "口袋里有一本七伤权谱立刻从" & to1 & "口袋偷走了<font color=red>七伤拳谱</font>立刻跑回家把七伤拳练成.."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"七伤拳",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
	
case 10
xunfaqi=sjjh_name & "你发现" & to1 & "口袋里有一本九阳神功立刻从" & to1 & "口袋偷走了<font color=red>九阳神功</font>立刻跑回家把七伤拳练成.."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"圣火令",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
