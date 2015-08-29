<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'传授♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_jhdj<jhdj_cs then
	Response.Write "<script language=JavaScript>{alert('提示：传内力最少需要["&jhdj_cs&"]级,才可以操作！');}</script>"
	Response.End
end if
if Weekday(date())=7 and (Hour(time())>=19 and Hour(time())<21) then
	Response.Write "<script language=JavaScript>{alert('提示：今天是夺宝时间，此功能禁用两个小时！');}</script>"
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
'对系统禁止字符处理
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【传授】<font color=" & saycolor & ">"+cuan(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'传授
function cuan(fn1,to1)
fn1=abs(fn1)
if fn1<200 then
	Response.Write "<script language=JavaScript>{alert('提示：传送内力一次最少200的，别太小气！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 等级,grade,内力 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if fn1>30000 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：传内力不要大于30000好不！');}</script>"
	Response.End
end if
if rs("内力")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的内力不足！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 内力=内力-" & fn1 & " where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 内力=内力+" & fn1 & " where 姓名='" & to1 &"'"
cuan= "##发功，满脸通红，呼呼直喘,头上青烟直冒，终于把" & fn1 & "的内力传给了%%了，%%万分感谢！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>