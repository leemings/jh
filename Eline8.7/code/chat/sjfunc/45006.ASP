<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'魔力钻石♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")

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
says="<font color=green>【魔力钻石】<font color=" & sayscolor & ">"+molishi(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)



function molishi(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w6,w9,法力,等级 FROM 用户 WHERE 姓名="&"'" & sjjh_name &"'",conn,3,3
fla=rs("法力")
dj=rs("等级")
w6w=rs("w9")
if fla<1000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得1000点啊！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if dj<20 then
Response.Write "<script language=JavaScript>{alert('此功能需要20级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

conn.execute "update 用户 set 法力=法力-1000 where 姓名="&"'" & sjjh_name &"'"
wuyn=iswp(w6w,"魔力钻石")
	if wuyn=1 then
	fq=abate(w6w,"魔力钻石",1)
	conn.execute "update 用户 set  w9='"&fq&"' where 姓名='"&sjjh_name&"'"
	else
	Response.Write "<script language=JavaScript>{alert('你有魔力钻石吗？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
	end if

	conn.execute "update 用户 set 法力=法力+5000 where 姓名="&"'" & sjjh_name &"'"
	molishi="##从口袋里拿出魔力钻石，轻轻一擦,魔力钻石金光闪闪，倾刻间<font color=red>##</font>的法力暴涨5000点." 
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>