<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'配制宝石
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
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
says="<font color=green>【魔法师·配制宝石】<font color=" & sayscolor & ">"+peibashi(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function peibashi(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,w9,等级,法力,职业 FROM 用户 WHERE 姓名="&"'" & aqjh_name &"'",conn,3,3
fali=rs("法力")
w6w=rs("w6")
if rs("等级")<50 then
Response.Write "<script language=JavaScript>{alert('此功能需要50级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("职业")<>"魔法师" then
	Response.Write "<script language=JavaScript>{alert('你只是普通子民呢！！请去职业转换为魔法师！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

if fali<3000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得3000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
wuyn=iswp(w6w,"红宝石")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('你没有红宝石？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
wuyn=iswp(w6w,"绿宝石")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('你没有绿宝石吗？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
wuyn=iswp(w6w,"蓝宝石")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('你没有蓝宝石吗？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select w6,w9 FROM 用户 WHERE 姓名="&"'" & aqjh_name &"'",conn,3,3
fq=abate(rs("w6"),"蓝宝石",1)
fq=abate(fq,"红宝石",1)
fq=abate(fq,"绿宝石",1)
conn.execute "update 用户 set  w6='"&fq&"' where 姓名='"&aqjh_name&"'"

fq1=add(rs("w9"),"魔力钻石",1)
conn.execute "update 用户 set  w9='"&fq1&"' where 姓名='"&aqjh_name&"'"
conn.execute "update 用户 set 法力=法力-3000 where 姓名="&"'" & aqjh_name &"'"
peibashi="##取出红、绿、蓝三种宝石，三种宝石结合在一起，一道光芒升起，<img src='img/look52.gif'>三种宝石化成了魔力钻石." 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>