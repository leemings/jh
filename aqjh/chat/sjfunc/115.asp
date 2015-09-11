<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'点石成金
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
says="<font color=red>【爱情魔法】<font color=" & saycolor & ">"+dscj()+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'点石成金
function dscj()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 金币,武功,等级,操作时间 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
dscj=DateDiff("n",rs("操作时间"),now())
if dscj<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=10-dscj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("武功")<200000  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：需要武功200000以上才可以操作！');}</script>"
	Response.End
end if
if rs("金币") then
	conn.execute "update 用户 set 金币=金币+2,武功=武功-200000,操作时间=now() where 姓名='" & aqjh_name &"'"
	dscj="##<bgsound src=wav/dscj.wav loop=1>使用【爱情魔法】点石成金,武功失去<font color=red>-200000</font>点,金币增加<font color=red>+5</font>个,##觉的爱情江湖实在太有意思了!"
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