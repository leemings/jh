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
r=int(rnd*51)+1
select case r
case 1
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>风之舞</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f4.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"风之舞",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 2
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠给您礼物[<b><font color=red>鹊桥相会</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f3.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"鹊桥相会",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 3
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>好梦</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f21.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"好梦",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 4
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠给您礼物[<b><font color=red>柔情蜜意</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f5.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"柔情蜜意",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 5
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠给您礼物[<b><font color=red>玫瑰衷情</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f7.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"玫瑰衷情",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 6
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠给您礼物[<b><font color=red>无限爱意</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f1.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"无限爱意",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 7
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠给您礼物[<b><font color=red>爱的韵味</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f2.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"风之舞",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 8
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>纯洁的爱</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f8.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"纯洁的爱",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 9
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>情深款款</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f16.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"情深款款",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 9
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>爱之泉</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f18.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"爱之泉",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 10
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>温情绽放</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f22.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"温情绽放",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 11
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>可人儿</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f19.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"可人儿",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 12
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>热恋</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f14.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"热恋",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 13
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>比翼双飞</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f13.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"比翼双飞",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 14
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>想你念你</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f9.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"想你念你",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 15
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>蝶之舞</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f12.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"蝶之舞",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 16
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>美满生活</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f24.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"美满生活",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 17
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>恋爱热吻</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f6.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"恋爱热吻",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 18
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>永远有你</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f20.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"永远有你",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 19
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>情有独钟</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f10.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"情有独钟",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 20
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>爱慕</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f11.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"爱慕",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 21
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>一生有你</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/f4.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"一生有你",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 22
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠给您礼物[<b><font color=red>浓情似火</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower17.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"浓情似火",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 23
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>百年好合</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower18.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"百年好合",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 24
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠给您礼物[<b><font color=red>美满幸福</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower19.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"美满幸福",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 25
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠给您礼物[<b><font color=red>情意浓浓</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower21.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"情意浓浓",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 26
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠给您礼物[<b><font color=red>健康永远</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower22.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"健康永远",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 27
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠给您礼物[<b><font color=red>一往情深</font></b>]<img src='picwords/2.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower33.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"一往情深",29)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 28
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>红袖留香</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower24.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"红袖留香",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 29
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>风情玫瑰</font></b>]<img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower25.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"风情玫瑰",9)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 30
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>特别的爱</font></b>]<img src='picwords/3.gif'><img src='picwords/1.gif'>朵<img src='../hcjs/jhjs/images/flower30.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"特别的爱",31)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 31
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>跳舞的玫瑰</font></b>]<img src='picwords/1.gif'><img src='picwords/1.gif'>朵<img src='../hcjs/jhjs/images/flower390.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"跳舞的玫瑰",11)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 32
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>拍手的玫瑰</font></b>]<img src='picwords/8.gif'>朵<img src='../hcjs/jhjs/images/flower350.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"拍手的玫瑰",8)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 33
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>失望的玫瑰</font></b>]<img src='picwords/7.gif'>朵<img src='../hcjs/jhjs/images/flower320.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"失望的玫瑰",7)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 34
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>想飞的玫瑰</font></b>]<img src='picwords/6.gif'>朵<img src='../hcjs/jhjs/images/flower400.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"想飞的玫瑰",6)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 35
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>摇摆的玫瑰</font></b>]<img src='picwords/5.gif'>朵<img src='../hcjs/jhjs/images/flower330.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"摇摆的玫瑰",5)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 36
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>转圈的玫瑰</font></b>]<img src='picwords/4.gif'>朵<img src='../hcjs/jhjs/images/flower340.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"转圈的玫瑰",4)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 37
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>致谢的玫瑰</font></b>]<img src='picwords/8.gif'><img src='picwords/8.gif'>朵<img src='../hcjs/jhjs/images/flower360.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"致谢的玫瑰",88)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 38
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>摇头的玫瑰</font></b>]<img src='picwords/8.gif'>朵<img src='../hcjs/jhjs/images/flower370.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"摇头的玫瑰",8)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 39
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>蹦跳的玫瑰</font></b>]<img src='picwords/7.gif'>朵<img src='../hcjs/jhjs/images/flower380.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"蹦跳的玫瑰",7)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 40
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>点头的玫瑰</font></b>]<img src='picwords/8.gif'>朵<img src='../hcjs/jhjs/images/flower310.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"点头的玫瑰",8)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 41
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>梦中情人</font></b>]<img src='picwords/1.gif'><img src='picwords/8.gif'>朵<img src='../hcjs/jhjs/images/flower48.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"梦中情人",18)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 42
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>爱到永远</font></b>]<img src='picwords/1.gif'><img src='picwords/9.gif'>朵<img src='../hcjs/jhjs/images/flower49.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"爱到永远",19)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 43
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>温馨的爱</font></b>]<img src='picwords/1.gif'><img src='picwords/7.gif'>朵<img src='../hcjs/jhjs/images/flower47.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"温馨的爱",17)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 44
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>此情可待</font></b>]<img src='picwords/1.gif'><img src='picwords/6.gif'>朵<img src='../hcjs/jhjs/images/flower45.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"此情可待",16)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 45
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>水晶之恋</font></b>]<img src='picwords/2.gif'><img src='picwords/0.gif'>朵<img src='../hcjs/jhjs/images/flower41.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"水晶之恋",20)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 46
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>爱的诉说</font></b>]<img src='picwords/2.gif'><img src='picwords/1.gif'>朵<img src='../hcjs/jhjs/images/flower40.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"爱的诉说",21)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 47
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>千里相思</font></b>]<img src='picwords/2.gif'><img src='picwords/8.gif'>朵<img src='../hcjs/jhjs/images/flower38.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"千里相思",28)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 48
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>此情不渝</font></b>]<img src='picwords/2.gif'><img src='picwords/6.gif'>朵<img src='../hcjs/jhjs/images/flower36.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"此情不渝",26)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 49
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>情定一生</font></b>]<img src='picwords/2.gif'><img src='picwords/7.gif'>朵<img src='../hcjs/jhjs/images/flower34.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"情定一生",27)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 50
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>我的牵挂</font></b>]<img src='picwords/1.gif'><img src='picwords/8.gif'>朵<img src='../hcjs/jhjs/images/flower37.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"我的牵挂",18)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 51
xunfaqi=aqjh_name & "在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>两情相悦</font></b>]<img src='picwords/3.gif'><img src='picwords/0.gif'>朵<img src='../hcjs/jhjs/images/flower32.gif'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"两情相悦",30)
	conn.execute "update 用户 set  w7='"&xianhua&"' where 姓名='"&aqjh_name&"'"
	rs.close

end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
