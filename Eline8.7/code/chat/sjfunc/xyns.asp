<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'新阴那山♀wWw.51eline.com♀
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
says="<font color=red>【一线魔法】<font color=" & saycolor & ">"+xyns()+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'新阴那山
function xyns()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select 武功,攻击,等级,操作时间 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
xyns=DateDiff("n",rs("操作时间"),now())
if xyns<(int(rnd*20)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*20)+1)-xyns
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
	conn.execute "update 用户 set 攻击=攻击+100,武功=武功-100000,操作时间=now() where 姓名='" & sjjh_name &"'"
	xyns="##<bgsound src=wav/xyns.wav loop=1>使用【一线魔法】新阴那山,去了阴那山以后武功失去<font color=red>-100000</font>点,攻击增加<font color=red>+100</font>点,##连声说下次我还要去阴那山!"
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