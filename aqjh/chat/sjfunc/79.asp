<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'送豆子
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_sq then
	Response.Write "<script language=JavaScript>{alert('提示：送豆子需要["&jhdj_sq&"]级才可以操作！');}</script>"
	Response.End
end if
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(0)<>"聊天大厅" then
 Response.Write "<script language=javascript>{alert('提示：送东西请去大厅！');}</script>"
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

'对暂离开、点哑穴判断
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
says="<bgsound src=wav/THANKS.wav loop=1><font color=green>【送豆子】<font color=" & saycolor & ">"+givedouzi(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'送豆子
function givedouzi(fn1,to1)
fn1=int(abs(fn1))
if (fn1<10 or fn1>500) and aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：送豆子最少10点，最多500点！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 泡豆点数,银两,知质 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("泡豆点数")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你有那么多的豆点吗？看好再说！');}</script>"
	Response.End
end if
if rs("银两")<100000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：赠送豆子需要10万两银子！');}</script>"
	Response.End
end if
if rs("知质")<40 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：赠送豆子需要40点知质！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 泡豆点数=泡豆点数-" & fn1 & ",银两=银两-100000,知质=知质-40 where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 泡豆点数=泡豆点数+" & fn1 & " where 姓名='" & to1 &"'"
givedouzi="##把" & fn1 & "豆点送给了%%[$$redb"&fn1&"$$b]点希望我们可以成为好朋友，有机会也送我一些好东西~~"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','操作','送豆点"&fn1&"个')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
