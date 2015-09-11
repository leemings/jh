<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'存水属性
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
says="<font color=green>【存水属性】<font color=" & saycolor & ">"+putfali(mid(says,i+1))+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'存水属性
function putfali(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 水加,结算日期,水 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
bankmoney=rs("水加")
lastdate=date()-rs("结算日期")
money=rs("水")
newbankmoney=int(bankmoney+bankmoney*0.0001*lastdate)
fn1=abs(fn1)
if fn1<=0 then
	if bankmoney=<0 then
		putfali="##你在水存放并没有水,金每天利息为:0.0001%,欢迎存放!"
	else
		putfali="##在水库存有:"& bankmoney &"点,在:"& rs("结算日期") &"存入,按0.0001%利,水库现在有:"& newbankmoney &"点水属性!"
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if



if money<fn1 then
	Response.Write "<script language=JavaScript>{alert('你哪里有那么多的水属性啊？？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
putfali="##在水属性宝瓶里放入了:"& fn1 &"点水属性,水属性宝瓶现存有水属性点:"& newbankmoney+fn1 &"点,继续存啊，保险啊!"
conn.execute "update 用户 set 水=水-"  & fn1 & ",水加="& newbankmoney+fn1 &",结算日期=date() where 姓名='" & aqjh_name &"'"

rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
