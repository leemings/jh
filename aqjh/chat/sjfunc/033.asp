<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'送金卡
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then 
	Response.Write "<script Language=Javascript>alert('提示：你不是正站长，你不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
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
says="<bgsound src=wav/THANKS.wav loop=1><font color=green>【站长送金卡】<font color=" & saycolor & ">"+givegold(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'送金卡
function givegold(fn1,to1)
fn1=int(abs(fn1))
if (fn1<10 or fn1>100) and aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：送金卡最少10个，最多100个！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 会员金卡,银两,轻功,grade FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("会员金卡")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你有那么多的会员金卡吗？看好再说！');}</script>"
	Response.End
end if
if rs("银两")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：赠送会员金卡需要1000两银子的出关费！');}</script>"
	Response.End
end if
if rs("轻功")<50000 then
Response.Write "<script language=JavaScript>{alert('提示：赠送会员金卡需要2000点轻功的手续费！！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 轻功=轻功-2000,会员金卡=会员金卡-" & fn1 & ",银两=银两-1000 where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 会员金卡=会员金卡+" & fn1 & " where 姓名='" & to1 &"'"
givegold="##把" & fn1 & "块金卡送给了%%[$$redb"&fn1&"$$b]个，此次赠送操作成功!"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','操作','送金卡"&fn1&"个')"
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
