<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc.asp"-->
<%'国家招收
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
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
says="<bgsound src=wav/go.wav loop=1><font color=green>【国家招人】<font color=" & saycolor & ">"+zsdz(mid(says,i+1))+"</font>"
call chatsay("对赌",towhoway,towho,saycolor,addwordcolor,addsays,says)

'对招收弟子点
function zsdz(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 职位,国家,grade from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if rs("grade")<3 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你不是君主也不是丞相,将军\n不可以招弟子！');}</script>"
	Response.End
end if
guo=rs("国家")
rs.close
set rs=nothing
conn.close
set conn=nothing
randomize()
regjm=int(rnd*3348998)+1
zsdz="[##]在江湖里下个★天下布告★,上面名曰：【"&guo&"】扩充国力,争霸天下，现在广招贤材!"&fn1&"欢迎兄弟们加入!<input  type=button value='加入!' onClick=javascript:zsdz"&regjm&".disabled=1;window.open('../../guojia/gj1.asp?id="&guo&"','d') name=zsdz"&regjm&">"
end function
%>
