<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'寻找卡片
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
	says=bdsays(says)
end if
f=Minute(time())
if f<1 or f>15 then
	Response.Write "<script language=JavaScript>{alert('永不放弃现在还没有发放卡片，寻找卡片时间为每小时的前15分钟！');window.close();}</script>"
	Response.End 
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【寻找卡片】<font color=" & saycolor & ">"+xunfaqi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'寻找卡片
function xunfaqi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 等级,会员等级,金币,操作时间,职业,转生 FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn,2,2
dj=rs("等级")
sj=DateDiff("n",rs("操作时间"),now())
if sj<3 then
	ss=3-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('请等上"&ss&"分再来寻找！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("职业")<>"冒险家" then
	Response.Write "<script language=JavaScript>{alert('你非冒险家，不能找宝物！！请去职业转换为冒险家！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
fla=rs("金币")
if rs("金币")<50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的金币不够无法施展呀，至少也得80个啊！');window.close();}</script>"
	response.end
end if
if dj<100 and rs("转生")<1 then
Response.Write "<script language=JavaScript>{alert('此功能需要[100]级战斗等级呀（转生人除外）！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
hydj=rs("会员等级")
if hydj<3 then
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：会员等级为3级才可以进行寻找卡片！购买加QQ51726805 ');}</script>"
	Response.End
end if

rs.close
conn.execute "update 用户 set 金币=金币-50,操作时间=now()  where 姓名='"&aqjh_name&"'"
randomize 
r=int(rnd*12)+1
select case r
case 1
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到大方的<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>福神卡</font></b>]<img src='picwords/1.gif'>张。"
	rs.open "SELECT w5 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"福神卡",1)
	conn.execute "update 用户 set  w5='"&kapian&"' where 姓名='"&aqjh_name&"'"
	rs.close


case 2
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到大方的<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>查税卡</font></b>]<img src='picwords/1.gif'>张。"
	rs.open "SELECT w5 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"查税卡",1)
	conn.execute "update 用户 set  w5='"&kapian&"' where 姓名='"&aqjh_name&"'"
	rs.close

	
case 3
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到大方的<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>涨钱卡</font></b>]<img src='picwords/1.gif'>张。"
	rs.open "SELECT w5 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"涨钱卡",1)
	conn.execute "update 用户 set  w5='"&kapian&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 4
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到大方的<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>练功卡</font></b>]<img src='picwords/1.gif'>张。"
	rs.open "SELECT w5 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"练功卡",1)
	conn.execute "update 用户 set  w5='"&kapian&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 5
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到大方的<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>亲亲卡</font></b>]<img src='picwords/2.gif'>张。"
	rs.open "SELECT w5 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"亲亲卡",2)
	conn.execute "update 用户 set  w5='"&kapian&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 6
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到大方的<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>种花卡</font></b>]<img src='picwords/2.gif'>张。"
	rs.open "SELECT w5 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"种花卡",2)
	conn.execute "update 用户 set  w5='"&kapian&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 7
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到大方的<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>健身卡</font></b>]<img src='picwords/1.gif'>张。"
	rs.open "SELECT w5 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"健身卡",1)
	conn.execute "update 用户 set  w5='"&kapian&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 8
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到大方的<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>抱抱卡</font></b>]<img src='picwords/2.gif'>张。"
	rs.open "SELECT w5 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"抱抱卡",2)
	conn.execute "update 用户 set  w5='"&kapian&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 9
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到大方的<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>财神卡</font></b>]<img src='picwords/1.gif'>张。"
	rs.open "SELECT w5 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"财神卡",1)
	conn.execute "update 用户 set  w5='"&kapian&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 10
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到大方的<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>养猪卡</font></b>]<img src='picwords/1.gif'>张。"
	rs.open "SELECT w5 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"养猪卡",1)
	conn.execute "update 用户 set  w5='"&kapian&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 11
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。本想乘此机会大捞一把，哪知刚进门就踩到狗大便，武功减少200，真是衰啊！~"
	rs.open "SELECT 武功 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	conn.execute "update 用户 set  武功=武功-200 where 姓名='"&aqjh_name&"'"
	rs.close
case 12
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。赶紧大步特步的跑过去，哪知碰到小石头，摔了一绞，体力减少200点，背到家了~！"
	rs.open "SELECT 体力 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	conn.execute "update 用户 set  体力=体力-200 where 姓名='"&aqjh_name&"'"
	rs.close
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
