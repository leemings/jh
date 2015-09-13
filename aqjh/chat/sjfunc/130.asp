<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'新阴那山
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
says="<font color=red>【快乐魔法】<font color=" & saycolor & ">"+xyns()+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'新阴那山
function xyns()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 武功,攻击,等级,操作时间 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
xyns=DateDiff("n",rs("操作时间"),now())
if xyns<5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=5-xyns
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("武功")<100000  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：需要武功100000以上才可以操作！');}</script>"
	Response.End
end if
if rs("攻击") then
	conn.execute "update 用户 set 攻击=攻击+100,武功=武功-100000,操作时间=now() where 姓名='" & aqjh_name &"'"
	xyns="##<bgsound src=wav/xyns.wav loop=1>使用【快乐魔法】新阴那山,去了阴那山以后武功失去<font color=red>-100000</font>点,攻击增加<font color=red>+100</font>点,##连声说下次我还要去雁南飞!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
