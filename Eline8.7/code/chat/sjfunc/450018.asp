<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'生日蛋糕♀wWw.happyjh.com♀
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
says="<font color=green>【生日蛋糕】<font color=" & sayscolor & ">"+shendangao(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)




function shendangao(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w6,法力,等级 FROM 用户 WHERE 姓名="&"'" & sjjh_name &"'",conn,3,3
fla=rs("法力")
dj=rs("等级")
w6w=rs("w6")
if fla<1000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得1000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj<10 then
Response.Write "<script language=JavaScript>{alert('此功能需要10级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 法力=法力-1000 where 姓名="&"'" & sjjh_name &"'"
fqsl=mywpsl(w6w,"生日蛋糕")
if fqsl<5 then
Response.Write "<script language=JavaScript>{alert('你有5个生日蛋糕吗？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
fq=abate(w6w,"生日蛋糕",5)
conn.execute "update 用户 set  w6='"&fq&"' where 姓名='"&sjjh_name&"'"
conn.execute "update 用户 set 体力=体力+500000 where 姓名="&"'" & sjjh_name &"'"
shendangao="##哼哼地从地上捡起一把雪饮狂刀把自己的肚子剖了开来塞了一块法器生日蛋糕<img src='pic/dz59.gif'>进去，嘿嘿，来吧，打我呀，有本事打死我啊，<font color=red>##</font>的体力暴涨500000点." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing

end function 
%>