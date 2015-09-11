<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'送钱
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
	Response.Write "<script language=JavaScript>{alert('提示：送钱需要["&jhdj_sq&"]级才可以操作！');}</script>"
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
says="<font color=green>【送钱】<font color=" & saycolor & ">"+give(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'送钱
function give(fn1,to1)
fn1=int(abs(fn1))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 银两 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if Instr(Application("aqjh_admin_send"),"|" & aqjh_name & "|")=0 then
if (fn1<100 or fn1>1000000) and aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：送钱最少100两，最多100万两！');}</script>"
	Response.End
end if
if rs("银两")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你有那么多的钱吗？看好再说！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 银两=银两-" & fn1 & " where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 银两=银两+" & int(fn1*0.8) & " where 姓名='" & to1 &"'"
give="##把" & fn1 & "两白银送给了%%,这下可把%%乐的直蹦,连声说谢谢,%%交纳个人所得税%20……<bgsound src=wav/THANKS.wav loop=1>"
else
if fn1<10000 then
give="##:你也太小气了吧，身为财神爷，就送这么一点？"
exit function
elseif fn1>20000000 then
give="##:就算你是财神爷，你也不能一出手就" & fn1 & "两白银啊！2000万最多了吧。你当金库银子不是你自已的不心疼啊？"
exit function
end if
conn.execute "update 用户 set 银两=银两+" &fn1& " where 姓名='" & to1 &"'"
give="%%可能是好事做的太多了，感动了财神爷<font color=red>["&aqjh_name&"]</font>天上掉下"&fn1&"两白银<img src='img/251.GIF'><bgsound src=wav/THANKS.wav loop=1>"
end if
conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','" & to1 & "','操作',now(),'赠送白银" & fn1 & "两')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
