<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'邀请加组
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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs1.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,1,3
zid=rs1("id")
rs1.close
rs1.open "select * FROM 组队 WHERE 组长=" & zid &"",conn,1,3
if rs1.eof and rs1.bof then
	Response.Write "<script language=JavaScript>{alert('提示：要邀请加入组需要建立组才可以操作！');}</script>"
	Response.End
rs1.close
end if
if towho=aqjh_name or towho=application("aqjh_automanname") then
	towho="大家"
else
	call dianzan(towho)
end if
if towho="大家" then
	Response.Write "<script language=JavaScript>{alert('提示：请所有人加入吗？组限制只有"&application("aqjh_zudui")&"个人呢！');}</script>"
	Response.End
end if
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=1
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<bgsound src=wav/go.wav loop=1><font color=green>【邀请加入组】<font color=" & saycolor & ">"+zsdz(mid(says,i+1))+"</font>"
call chatsay("加组",towhoway,towho,saycolor,addwordcolor,addsays,says)

function zsdz(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 组队 where 组长="&zid&"",conn,1,3
if rs.eof and rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你不是组长\n不可以邀请别人加入！');}</script>"
	Response.End
end if
zu=rs("组名")
rs.close
set rs=nothing
conn.close
set conn=nothing
randomize()
regjm=int(rnd*3348998)+1
zsdz="[##]对<font color=blue>"&towho&"</font>喊到,我们组["&zu&"]人多,力量大,现在广招组成员!"&fn1&"欢迎兄弟们加入!<input  type=button value='加入!' onClick=javascript:zsdz"&regjm&".disabled=1;window.open('zhu/join.asp?name1="&towho&"&name2="&aqjh_name&"') name=zsdz"&regjm&">"
end function
%>
