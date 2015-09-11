<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'送金币
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
	Response.Write "<script language=JavaScript>{alert('提示：送金币需要["&jhdj_sq&"]级才可以操作！');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(0)<>"聊天大厅" then
 Response.Write "<script language=javascript>{alert('提示：送金笔请去大厅，这里没人在！');}</script>"
 Response.End
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
says="<bgsound src=wav/THANKS.wav loop=1><font color=green>【送金币】<font color=" & saycolor & ">"+givegold(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'送金币
function givegold(fn1,to1)
fn1=int(abs(fn1))
if (fn1<1 or fn1>100) and aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：送金币最少1个，最多100个！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 金币,银两,轻功,grade FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("金币")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你有那么多的金币吗？看好再说！');}</script>"
	Response.End
end if
if rs("银两")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：赠送金币需要1000两银子的出关费！');}</script>"
	Response.End
end if
if rs("grade")<4 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：只有掌门跟管理员才可以送金币！');}</script>"
	Response.End
end if
if rs("轻功")<2000 then
Response.Write "<script language=JavaScript>{alert('提示：赠送金币需要2000点轻功的手续费！！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 轻功=轻功-2000,金币=金币-" & fn1 & ",银两=银两-1000 where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 金币=金币+" & fn1 & " where 姓名='" & to1 &"'"
givegold="##把" & fn1 & "金币送给了%%[$$redb"&fn1&"$$b]个，此次赠送操作成功!"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','操作','送金币"&fn1&"个')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
onn.close
set conn=nothing
end function
%>