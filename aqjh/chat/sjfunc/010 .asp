<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'寻找鲜花
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

act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【寻找鲜花】<font color=" & saycolor & ">"+xunfaqi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'寻找鲜花
function xunfaqi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn,2,2
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
f=Minute(time())
if f<25 or f>50 then
	Response.Write "<script language=JavaScript>{alert('花店老板现在没有开放花库，寻找鲜花时间为每小时的25-50分钟！');window.close();}</script>"
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
if rs("金币")<4 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的金币不够无法施展呀，至少也得4个啊！');window.close();}</script>"
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

rs.close
conn.execute "update 用户 set 金币=金币-4,操作时间=now() where 姓名='"&aqjh_name&"'"
randomize 
r=int(rnd*17)+1
select case r
case 1
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到<font color=#0000FF>花店老板</font>赠送给您礼物[<b><font color=red>玫瑰</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower19.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"玫瑰",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close


case 2
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了江湖MM也在那里聚会。收到MM<font color=#0000FF>月光</font>赠给您礼物[<b><font color=red>白玫瑰</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower20.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"白玫瑰",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 3
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到<font color=#0000FF>花店老板</font>赠送给您礼物[<b><font color=red>好梦</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower21.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"好梦",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 4
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到<font color=#0000FF>花店老板</font>赠给您礼物[<b><font color=red>红袖留香</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower22.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"红袖留香",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 5
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到<font color=#0000FF>花店老板</font>赠给您礼物[<b><font color=red>醉百合</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower23.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"醉百合",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 6
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到<font color=#0000FF>花店老板</font>赠给您礼物[<b><font color=red>昨日香</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower24.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"昨日香",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 7
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到<font color=#0000FF>花店老板</font>赠给您礼物[<b><font color=red>迎风笑</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower25.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"迎风笑",99)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 8
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到<font color=#0000FF>花店老板</font>赠送给您礼物[<b><font color=red>月季</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower26.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"月季",99)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 9
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到<font color=#0000FF>花店老板</font>赠送给您礼物[<b><font color=red>蝴蝶兰</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower27.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"蝴蝶兰",99)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 10
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到好心的<font color=#0000FF>花店老板</font>赠送给您礼物[<b><font color=red>浓情似火</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower28.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"浓情似火",99)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 11
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到好心的<font color=#0000FF>花店老板</font>赠送给您礼物[<b><font color=red>雪莲花</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower29.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"雪莲花",99)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 12
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到好心的<font color=#0000FF>花店老板</font>赠送给您礼物[<b><font color=red>郁金香</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower30.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"郁金香",99)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 13
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到好心的<font color=#0000FF>花店老板</font>赠送给您礼物[<b><font color=red>爱的诉说</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower31.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"爱的诉说",99)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 14
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到好心的<font color=#0000FF>花店老板</font>赠送给您礼物[<b><font color=red>满天星</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower32.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"满天星",99)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 15
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到好心的<font color=#0000FF>花店老板</font>赠送给您礼物[<b><font color=red>示爱玫瑰</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower33.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"示爱玫瑰",99)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 16
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到好心的<font color=#0000FF>花店老板</font>赠送给您礼物[<b><font color=red>爱你永不变</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower34.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"爱你永不变",99)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 17
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了花店老板也在那里聚会。收到好心的<font color=#0000FF>花店老板</font>赠送给您礼物[<b><font color=red>雏菊</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower35.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"雏菊",99)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
		
		
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
