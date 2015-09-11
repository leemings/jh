<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'闭目养神
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
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
says="<font color=red>【闭目】<font color=" & saycolor & ">"+bimu()+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'闭目养神
function bimu()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 内力,体力,等级,体力加,操作时间 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("内力")<1100  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：需要体力1000才可以操作！');}</script>"
	Response.End
end if
if rs("体力")<rs("等级")*aqjh_tlsx+5260+rs("体力加") then
	conn.execute "update 用户 set 体力=体力+3500,内力=内力-1000,操作时间=now() where 姓名='" & aqjh_name &"'"
	bimu="##<bgsound src=wav/dz.wav loop=1>进行练功打坐,内力失去<font color=red>-1000</font>点,体力提升<font color=red>+1500</font>点,有得必有失嘛!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("体力加")<=rs("等级")*1800 then
	conn.execute "update 用户 set 体力加=体力加+9,内力=内力-1000,操作时间=now() where 姓名='" & aqjh_name &"'"
	bimu="##<bgsound src=wav/dz.wav loop=1>进行练功打坐,内力失去-1000体力上限提升+9,有得必有失嘛!"
else
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：现在你的上限满了，等升了级再练吧');}</script>"
	Response.End
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
