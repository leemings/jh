<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'养心大法
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
says="<font color=red>【快乐魔法】<font color=" & saycolor & ">"+yxdf()+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'养心大法
function yxdf()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 体力,内力,等级,操作时间 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
yxdf=DateDiff("n",rs("操作时间"),now())
if yxdf<5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=5-yxdf
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("体力")<10000  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：需要体力10000以上才可以操作！');}</script>"
	Response.End
end if
if rs("内力") then
	conn.execute "update 用户 set 内力=内力+1000,体力=体力-10000,操作时间=now() where 姓名='" & aqjh_name &"'"
	yxdf="##<bgsound src=wav/yxdf.wav loop=1>使用【快乐魔法】养心大法,体力失去<font color=red>-200000</font>点,内力增加<font color=red>+100000</font>点,以后不用买那么多药了！!"
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