<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<%'提出轻功
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
	Response.Write "<script language=JavaScript>{alert('提示：现在为竞标时间，不可取轻功操作!');window.close();}</script>"
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
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【提出轻功】<font color=" & saycolor & ">"+lift(mid(says,i+1))+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'提出轻功
function lift(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select 轻功加,结算日期,轻功 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
bankmoney=rs("轻功加")
lastdate=date()-rs("结算日期")
money=rs("轻功")
newbankmoney=int(bankmoney+bankmoney*0.0001*lastdate)
fn1=abs(fn1)
if fn1<=0 then
	if bankmoney=<0 then
		lift="##你在轻功宝塔并没有存放轻功点,官府代为保管轻功宝塔每天利息为:0.0001%,欢迎存取"
	else
		lift="##在轻功存有:"& bankmoney &"轻功,在:"& rs("结算日期") &"存入,按0.0001%利,轻功现在有:"& newbankmoney &"点轻功!"
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if newbankmoney<fn1 then
	Response.Write "<script language=JavaScript>{alert('提示：你哪里有那么多的轻功？？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
lift="##从轻功宝塔里取出:"& fn1 &"点轻功,轻功宝塔现存有轻功:"& newbankmoney-fn1 &"点,好好使用,别被别人吸了!"
conn.execute "update 用户 set 轻功=轻功+"  & fn1 & ",轻功加="& newbankmoney-fn1 &",结算日期=date() where 姓名='" & sjjh_name &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
