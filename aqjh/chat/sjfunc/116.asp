<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'新人费
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_grade<7 then
	Response.Write "<script language=JavaScript>{alert('提示：送钱需要["&aqjh_grade&"]级才可以操作！');}</script>"
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
says="<bgsound src=wav/THANKS.wav loop=1><font color=green>【新人费】<font color=" & saycolor & ">"+give(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'新人费
function give(fn1,to1)
fn1=int(abs(fn1))
if (fn1<100 or fn1>1000000) and aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：送钱最少100两，最多100万两！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 银两 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("银两")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你有那么多的钱吗？看好再说！');}</script>"
	Response.End
end if
rs.close
rs.open "select 魅力 FROM 用户 WHERE 姓名='" & to1 &"'",conn
if rs("魅力")>1000 then
Response.Write "<script language=JavaScript>{alert('谢谢你，但我已经不需要新人费帮助了!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 银两=银两-" & fn1 & " where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 道德=道德+1000 where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 魅力=魅力+2000,银两=银两+" & int(fn1*0.8) & " where 姓名='" & to1 &"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','奖励','新人费:送白银"&fn1&"两')"
give="##把" & fn1 & "两新人费送给了%%,这下可把%%乐的直蹦,连声说谢谢，好心有好报，##道德上涨<font color=red>1000</font>点，看得各位大虾眼红红，齐说：下次我一定比你先照顾新人……"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
