<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc.asp"-->
<%'招收弟子♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
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
says="<font color=green>【招收弟子】<font color=" & saycolor & ">"+zsdz(mid(says,i+1))+"</font>"
call chatsay("对赌",towhoway,towho,saycolor,addwordcolor,addsays,says)

'对招收弟子点
function zsdz(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 操作时间,身份,门派,grade from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<(int(rnd*3)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*3)+1)-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分再操作！');}</script>"
	Response.End
end if
if rs("grade")<4 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你不是掌门也不是长老\n不可以招弟子！');}</script>"
	Response.End
end if
mp=rs("门派")
conn.execute "update 用户 set 操作时间=now() where 姓名='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
randomize()
regjm=int(rnd*3348998)+1
zsdz="<bgsound src=wav/go.WAV loop=1>[##]对聊天室的人喊到,我们门派["&mp&"]人多,力量大,现在广招弟子!"&fn1&"欢迎兄弟们加入!<input  type=button value='加入!' onClick=javascript:zsdz"&regjm&".disabled=1;window.open('../jhmp/mp1.asp?id="&mp&"','d') name=zsdz"&regjm&">"
end function
%>
