<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'取钱♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if Weekday(date())=1 and Hour(time())>=20 and Hour(time())<21 then
	Response.Write "<script language=JavaScript>{alert('提示：现在为竞标时间，不可取钱操作!');window.close();}</script>"
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
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【取钱】<font color=" & saycolor & ">"+getyin(mid(says,i+1))+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'取钱
function getyin(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select 存款,结算日期,银两 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
bankmoney=rs("存款")
lastdate=date()-rs("结算日期")
money=rs("银两")
newbankmoney=int(bankmoney+bankmoney*0.0001*lastdate)
fn1=abs(fn1)
if fn1<=0 then
	if bankmoney=<0 then
		getyin="##你在江湖钱庄并没有存款,钱庄每天利息为:0.0001%,欢迎存取!"
	else
		getyin="##在钱庄存有:"& bankmoney &"两,在:"& rs("结算日期") &"存入,按0.0001%利,银行现在有:"& newbankmoney &"两!"
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if newbankmoney<fn1 then
	Response.Write "<script language=JavaScript>{alert('提示：你哪里有那么多的钱？？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
getyin="##向钱庄取钱:"& fn1 &"两,钱庄现有银两:"& newbankmoney-fn1 &"两,拿好你的钱,别被抢了!"
conn.execute "update 用户 set 银两=银两+"  & fn1 & ",存款="& newbankmoney-fn1 &",结算日期=date() where 姓名='" & sjjh_name &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>