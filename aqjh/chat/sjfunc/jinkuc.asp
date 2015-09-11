<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'存金币
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
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
says="<font color=green>【金币存库】<font color=" & saycolor & ">"+getyin(mid(says,i+1))+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'存金币
function getyin(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 金库,结算日期,金币 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
bankmoney=rs("金库")
lastdate=date()-rs("结算日期")
money=rs("金币")
newbankmoney=int(bankmoney+bankmoney*0.00001*lastdate)
fn1=abs(fn1)
if fn1<=0 then
	if bankmoney=<0 then
		getyin="##你在江湖金库并没有金币,金库每天利息为:0.00001%,欢迎存取!"
	else
		getyin="##在金库存有:"& bankmoney &"金币,在:"& rs("结算日期") &"存入,按0.00001%利,金库现在有:"& newbankmoney &"金币!"
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if money<fn1 then
	Response.Write "<script language=JavaScript>{alert('提示：你的金库哪有那么多金币？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
getyin="##向金库存币:"& fn1 &"个,金库现有金币:"& newbankmoney+fn1 &"个,欢迎存取!"
conn.execute "update 用户 set 金币=金币-"  & fn1 & ",金库="& newbankmoney+fn1 &",结算日期=date() where 姓名='" & aqjh_name &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>