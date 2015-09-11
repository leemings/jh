<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'魔力钻石
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
says="<font color=green>【魔力钻石】<font color=" & sayscolor & ">"+molishi(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function molishi(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,w9,法力,等级,智力,职业 FROM 用户 WHERE 姓名="&"'" & aqjh_name &"'",conn,3,3
fla=rs("法力")
dj=rs("等级")
w6w=rs("w9")
if fla<2000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得2000点啊！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("职业")<>"炼金师" then
	Response.Write "<script language=JavaScript>{alert('你只是普通子民呢！！请去职业转换为炼金师！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("智力")<2 then
Response.Write "<script language=JavaScript>{alert('你的智力不够无法施展呀，至少也得2点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj<55 then
Response.Write "<script language=JavaScript>{alert('此功能需要55级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

conn.execute "update 用户 set 法力=法力-2000,智力=智力-2 where 姓名="&"'" & aqjh_name &"'"
wuyn=iswp(w6w,"魔力钻石")
	if wuyn=1 then
	fq=abate(w6w,"魔力钻石",1)
	conn.execute "update 用户 set  w9='"&fq&"' where 姓名='"&aqjh_name&"'"
	else
	Response.Write "<script language=JavaScript>{alert('你有魔力钻石吗？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
	end if

	conn.execute "update 用户 set 法力=法力+10000,智力=智力+5 where 姓名="&"'" & aqjh_name &"'"
	molishi="##从口袋里拿出魔力钻石，轻轻一擦,魔力钻石金光闪闪，倾刻间<font color=red>##</font>的法力暴涨10000点，智力值上涨<font color=red>5</font>点." 
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>icwords/0.gif'>点." 
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>