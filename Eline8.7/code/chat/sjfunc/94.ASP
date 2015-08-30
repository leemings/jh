<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'送银币♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_jhdj<jhdj_sq then
	Response.Write "<script language=JavaScript>{alert('提示：送银币需要["&jhdj_sq&"]级才可以操作！');}</script>"
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

'对暂离开、点哑穴判断
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
says="<font color=green>【送银币】<font color=" & saycolor & ">"+givegold(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'送银币
function givegold(fn1,to1)
fn1=int(abs(fn1))
if (fn1<10 or fn1>100) and sjjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：送银币最少10个，最多100个！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 银币,银两 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if rs("银币")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你有那么多的银币吗？看好再说！');}</script>"
	Response.End
end if
if rs("银两")<1000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：赠送银币需要100万两银子！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 银币=银币-" & fn1 & ",银两=银两-1000000 where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 银币=银币+" & fn1 & " where 姓名='" & to1 &"'"
givegold="<bgsound src=wav/thanks.WAV loop=1>##把" & fn1 & "银币送给了%%[$$redb"&fn1&"$$b]个，此次赠送操作成功!"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& to1 &"','操作','送银币"&fn1&"个')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
