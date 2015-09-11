<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'传授体力
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_cs then
	Response.Write "<script language=JavaScript>{alert('提示：传授体力最少需要["&jhdj_cs&"]级,才可以操作！');}</script>"
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
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【传授体力】<font color=" & saycolor & ">"+cuan(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'传授体力
function cuan(fn1,to1)
fn1=abs(fn1)
if fn1<1000 then
	Response.Write "<script language=JavaScript>{alert('提示：传授体力一次最少1000的，别太小气！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 等级,grade,体力 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if fn1>100000 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：传授体力不要大于100000好不！');}</script>"
	Response.End
end if
if rs("体力")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的体力不足！');}</script>"
	Response.End
end if
rs.close
rs.open "select 体力 FROM 用户 WHERE 姓名='" & to1 &"'",conn
if rs("体力")>2000000 then
Response.Write "<script language=JavaScript>{alert('失败，他体力已经大于200万了，根本不需要你保护了!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 体力=体力-" & fn1 & " where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 体力=体力+" & fn1 & " where 姓名='" & to1 &"'"
cuan= "##发功，满脸通红，呼呼直喘,头上青烟直冒，终于把" & fn1 & "的体力传给了%%了，%%万分感谢！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>othing	
conn.close
set conn=nothing
end function
%>