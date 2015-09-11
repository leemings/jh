<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'乾坤一掷
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
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
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
rs.open "select 银两,操作时间,grade,保护,门派 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<2 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=2-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是出家人不能操作，我K你跟你说！');}</script>"
	Response.End
end if

if rs("银两")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你有那么多的钱吗？看好再说！');}</script>"
	Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是官府人员!');}</script>"
	Response.End
end if
if rs("保护")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你正在保护中!');}</script>"
	Response.End
end if
rs.close
rs.open "select 体力,门派,grade,保护,通缉 FROM 用户 WHERE 姓名='" & to1 &"'",conn,2,2
if rs("门派")="出家" and aqjh_grade<10 and rs("通缉")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：不可攻击出家人!');}</script>"
	Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：不可攻击官府人员!');}</script>"
	Response.End
end if
if rs("保护")=True and aqjh_grade<10 and rs("通缉")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方正在保护中!');}</script>"
	Response.End
end if
if rs("体力")<=1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：["&to1&"]的体力很虚弱了，你不要再攻击了!');}</script>"
	Response.End
end if
randomize ()
m1 = Int(100 * Rnd)+100
gjtl=int(fn1/m1)
gjtl=rs("体力")-gjtl
if gjtl<1000 then 	gjtl=1000
conn.execute "update 用户 set 体力="&gjtl&" where 姓名='" & to1 &"'"
conn.execute "update 用户 set 银两=银两-" & fn1 & ",操作时间=now() where 姓名='" & aqjh_name &"'"
moneygj="##从口袋中拿出白银" & fn1 & "两，使用了一招林月如的乾坤一掷,所有的银两是一种无形的力量，打得%%晕头转向，%%体力只剩下[$$redb"&gjtl&"$$b]点~~"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
