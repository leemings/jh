<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'公告♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_grade<grade_gg then
	Response.Write "<script language=JavaScript>{alert('提示：公告需要["&grade_gg&"]级才可以操作！');}</script>"
	Response.End
end if
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
says=gong(mid(says,i+1))
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'公告
function gong(fn1)
fn1=trim(fn1)
gong="<bgsound src=wav/gonggao.wav loop=1><table width=85% cellspacing=0 cellpadding=0 bgcolor=#009933 align=center bordercolorlight=000000 bordercolordark=FFFFFF border=1><tr><td align=center><font color=FFFFFF>★ 官 府 公 告 ★</font></td><tr><td bgcolor=CCCCFF> <div align='center'><font size=-1>"&fn1&"(管理员)</font></div></td></tr></table>"
end function
%>
