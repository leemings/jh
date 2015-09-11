<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'奖励
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_grade<4 then
	Response.Write "<script language=JavaScript>{alert('提示：想作什么呀，你的管理等级可不够呀！');}</script>"
	Response.End
end if
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
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【奖励】<font color=" & saycolor & ">"+jiangli(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'奖励
function jiangli(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,grade,身份 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
menpai=rs("门派")
if rs("grade")<4 and (rs("身份")<>"长老" or rs("身份")<>"掌门") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：想作什么呀，你的管理等级可不够呀！');}</script>"
	Response.End
end if
fn1=int(abs(fn1))
if fn1>100000 or fn1<100  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：想作什么呀，你的管理等级可不够呀！');}</script>"
	Response.End
end if
rs.close
rs.open "select h FROM p WHERE a='" & menpai & "'",conn,2,2
if rs("h")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：帮派里只有："&rs("h")&"两,哪里有那么多的钱呀！');}</script>"
	Response.End
end if
rs.close
rs.open "select 门派,门派基金,等级 FROM 用户 WHERE 姓名='" & to1 &"'",conn,2,2
if trim(rs("门派"))<> trim(menpai) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：搞错了吧，["&to1&"]并不是你派弟子！');}</script>"
	Response.End
end if
if rs("门派基金")<-(rs("等级")*20000) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：不行了呀，["&to1&"]已经拿了太多的钱！');}</script>"
	Response.End
end if
conn.execute "update p set h=h-" & fn1 & " where a='" & menpai & "'"
conn.execute "update 用户 set 银两=银两+" & fn1 & ",门派基金=门派基金-"& fn1 &" where 姓名='" & to1 &"'"
jiangli="##把自己门派["&menpai&"]的基金发给弟子%%,"& fn1 & "两作为奖励！%%乐的直蹦,连声说谢谢!"
fn1="门派："&menpai&"  银两："&fn1
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','奖励','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>