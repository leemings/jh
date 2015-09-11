<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'传送法力
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_so then
	Response.Write "<script language=JavaScript>{alert('提示：传送法力需要["&jhdj_so&"]级才可以操作！');}</script>"
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
says="<font color=green>【传送法力】<font color=" & saycolor & ">"+songfali(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'传送法力
function songfali(fn1,to1)
fn1=int(abs(fn1))
if (fn1<100 or fn1>500000) and aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：传送法力最少100点，最多5万点！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 法力,轻功,智力 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("法力")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你有那么多的法力吗？看好再说！');}</script>"
	Response.End
end if
if rs("轻功")<50000 then
Response.Write "<script language=JavaScript>{alert('你的轻功不够无法施展呀，至少也得5万点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("智力")<10 then
Response.Write "<script language=JavaScript>{alert('你的智力不够无法施展呀，至少也得10点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 智力=智力-10,轻功=轻功-2000,法力=法力-"& fn1 &" where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 法力=法力+"& fn1 &" where 姓名='" & to1 &"'"
songfali=aqjh_name & "把自己的"& fn1 &"点法力传送给了<font color=red>" & to1 & "</font>，<font color=red>" & to1 & "</font>法力大增，自己的智力减少了10点,轻功减少2000！"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','传授','传授法力"&fn1&"点')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>



