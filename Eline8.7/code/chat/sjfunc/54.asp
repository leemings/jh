<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc.asp"-->
<%'分手♀wWw.51eline.com♀
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
says="<font color=green>【情人分手】<font color=" & saycolor & ">"+fen(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'情人分手
function fen(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 会员金卡,情人,存款 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
qingren=rs("情人")
if rs("情人")="无" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你现在还没有情人不能分手！');}</script>"
	Response.End
end if
if rs("存款")<2000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：存款不够200万，不给200万分手费谁干！');}</script>"
	Response.End
end if
if rs("会员金卡")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：要想与情人分手需要1元会员金卡！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 存款=存款-2000000,会员金卡=会员金卡-1,情人='无' where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 存款=存款+2000000,情人='无' where 姓名='"&qingren&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
fen="[##]给情人{"&qingren&"}200万的分手费，与{"&qingren&"}分手了……"&fn1
end function
%>
