<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'生日蛋糕
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
says="<font color=green>【生日蛋糕】<font color=" & sayscolor & ">"+shendangao(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function shendangao(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,法力,等级,智力,职业,allvalue FROM 用户 WHERE 姓名="&"'" & aqjh_name &"'",conn,3,3
fla=rs("法力")
dj=rs("等级")
w6w=rs("w6")
if fla<5000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得5000点啊！');}</script>"
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
if rs("智力")<20 then
Response.Write "<script language=JavaScript>{alert('你的智力不够无法施展呀，至少也得20点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj<30 then
Response.Write "<script language=JavaScript>{alert('此功能需要30级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 法力=法力-5000,智力=智力-20 where 姓名="&"'" & aqjh_name &"'"
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
conn.execute "update 用户 set  w6='"&fq&"' where 姓名='"&aqjh_name&"'"
conn.execute "update 用户 set 体力=体力+2000000,allvalue=allvalue+10 where 姓名="&"'" & aqjh_name &"'"
shendangao="##哼哼地从地上捡起一把雪饮狂刀把自己的肚子剖了开来塞了一块生日蛋糕<img src='pic/dz59.gif'>进去，嘿嘿，来吧，打我呀，有本事打死我啊，<font color=red>##</font>的体力暴涨2000000点，跟着战斗经验增加，总积分上涨<font color=red>10</font>点." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing

end function 
%>