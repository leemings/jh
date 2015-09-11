<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'寻找会员金卡
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
if f<25 or f>40 then
	Response.Write "<script language=JavaScript>{alert('站长现在没有发放会员金卡，寻找金卡时间为每小时的25-40分钟！');window.close();}</script>"
	Response.End 
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【寻找会员金卡】<font color=" & saycolor & ">"+xunshuijing(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'寻找会员金卡
function xunshuijing(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 金币,会员等级,操作时间,转生,职业,等级 FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn,2,2
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
dj=rs("等级")
if dj<120 and rs("转生")<1 then
Response.Write "<script language=JavaScript>{alert('此功能需要[120]级战斗等级呀！（转生人除外）');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
fla=rs("金币")
if rs("金币")<6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的金币不够无法施展呀，至少也得6个啊！');window.close();}</script>"
	response.end
end if
hydj=rs("会员等级")
if hydj<2 then
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：会员等级为2级才可以进行寻找会员金卡!');}</script>"
	Response.End
end if
rs.close
conn.execute "update 用户 set 金币=金币-6,操作时间=now()  where 姓名='"&aqjh_name&"'"
randomize 
r=int(rnd*6)+1
select case r
case 1
xunshuijing=aqjh_name & "你寻遍了江湖各个角落终于找到了[<b><font color=red><img src='picwords/1.gif'></font></b>]张会员金卡<img src='card/jk.gif'>," & aqjh_name & "真是幸运呀!" 
	conn.execute "update 用户 set 会员金卡=会员金卡+1 where 姓名='" & aqjh_name &"'"
	

case 2
xunshuijing=aqjh_name & "你寻遍了江湖各个角落终于找到了[<b><font color=red><img src='picwords/2.gif'></font></b>]张会员金卡<img src='card/jk.gif'>," & aqjh_name & "真是幸运呀!" 
	conn.execute "update 用户 set 会员金卡=会员金卡+2 where 姓名='" & aqjh_name &"'"
	
case 3
xunshuijing=aqjh_name & "真幸运，你寻遍了江湖各个角落找到了[<b><font color=red><img src='picwords/5.gif'></font></b>]个金币" & aqjh_name & "增加了金币[<b><font color=red><img src='picwords/5.gif'></font></b>]个<img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'>!" 
	conn.execute "update 用户 set 金币=金币+5 where 姓名='" & aqjh_name &"'"

case 4
xunshuijing=aqjh_name & "真幸运，你寻遍了江湖各个角落找到了[<b><font color=red><img src='picwords/5.gif'></font></b>]个金币" & dljh_name & "增加了金币[<b><font color=red><img src='picwords/5.gif'></font></b>]个<img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'>!" 
	conn.execute "update 用户 set 金币=金币+5 where 姓名='" & aqjh_name &"'"

case 5
xunshuijing=aqjh_name & "好可惜，你寻遍了江湖各个角落找到了[<b><font color=red><img src='picwords/5.gif'></font></b>]个金币" & dljh_name & "增加了金币[<b><font color=red><img src='picwords/5.gif'></font></b>]个<img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'>!" 
	conn.execute "update 用户 set 金币=金币+5 where 姓名='" & aqjh_name &"'"

case 6
xunshuijing=aqjh_name & "你寻遍了江湖各个角落终于找到了[<b><font color=red><img src='picwords/5.gif'></font></b>]张会员金卡<img src='card/jk.gif'>," & aqjh_name & "损耗[<b><font color=red><img src='picwords/3.gif'><img src='picwords/0.gif'></font></b>]个金币!" 
	conn.execute "update 用户 set 金币=金币-30,会员金卡=会员金卡+5 where 姓名='" & aqjh_name &"'"
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
