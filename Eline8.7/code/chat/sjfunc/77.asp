<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'休身养性♀wWw.happyjh.com♀
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
says="<font color=red>【休身养性】<font color=" & saycolor & ">"+xsyx()+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'休身养性
function xsyx()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select 操作时间,泡豆点数,银两 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<(int(rnd*3)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*3)+1)-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("泡豆点数")<300 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的豆点数不够300不可以操作！');}</script>"
	Response.End
end if
if rs("银两")<1000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的银两不够100万不可以操作！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 武功加=武功加+150,体力加=体力加+150,内力加=内力加+150,泡豆点数=泡豆点数-300,银两=银两-1000000,操作时间=now() where 姓名='" & sjjh_name &"'"
xsyx="##<bgsound src=wav/dz.wav loop=1>进行休身养性的训练,消耗银两<font color=red>100</font>万,豆点<font color=red>300</font>点~~内力、体力、武功上限提高<font color=blue><b>150</b></font>点，努力的結果总是让人出乎意料~~"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>