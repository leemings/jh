<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'取水属性
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if Weekday(date())=1 and Hour(time())>=20 and Hour(time())<21 then
	Response.Write "<script language=JavaScript>{alert('提示：现在为竞标时间，不可取水属性操作!');window.close();}</script>"
	Response.End 
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
says="<font color=green>【取水属性】<font color=" & saycolor & ">"+getfali(mid(says,i+1))+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'取水属性
function getfali(fn1)
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
		getfali="##你在水宝瓶并没有存放水点,官府代为保管水宝瓶每天利息为:0.0001%,欢迎存取"
	else
		getfali="##在水存有:"& bankmoney &"水,在:"& rs("结算日期") &"存入,按0.0001%利,水现在有:"& newbankmoney &"点水!"
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if newbankmoney<fn1 then
	Response.Write "<script language=JavaScript>{alert('提示：你哪里有那么多的水？？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
getfali="##从水宝瓶里取出:"& fn1 &"点水,水宝瓶现存有法力:"& newbankmoney-fn1 &"点,好好使用,别被别人吸了!"
conn.execute "update 用户 set 水=水+"  & fn1 & ",水加="& newbankmoney-fn1 &",结算日期=date() where 姓名='" & aqjh_name &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
