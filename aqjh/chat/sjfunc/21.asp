<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'公告
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_grade<grade_gg then
	Response.Write "<script language=JavaScript>{alert('提示：公告需要["&grade_gg&"]级才可以操作！');}</script>"
	Response.End
end if
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
says=gong(mid(says,i+1))
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'公告
function gong(fn1)
fn1=trim(fn1)
gong="<bgsound src=wav/luo.wav loop=3><table align=center><td><table border=0 cellpadding=0 cellspacing=0 ><tr><td><img src=../chat/f2/img/rightct3.gif></td><td background=../chat/f2/img/rightct4.gif></td><td><img  src=../chat/f2/img/rightct1.gif></td></tr><tr><td background=../chat/f2/img/rightct8.gif ></td><td valign=center align=center><table align=center border=0 cellpadding=1 cellspacing=0 ><tr><td valign=center align=center><font style='font-size:9pt;color:red'>========官府公告========</font></td></tr><tr><td valign=center align=center><font style='font-size:9pt;color:green'>"&fn1&"(##)</td></tr></table></td><td background=../chat/f2/img/rightct08.gif></td></tr><tr><td><img src=../chat/f2/img/rightct9.gif></td><td background=../chat/f2/img/rightct10.gif></td><td><img src=../chat/f2/img/rightct11.gif></td></tr></table></td></table>"
end function
%>