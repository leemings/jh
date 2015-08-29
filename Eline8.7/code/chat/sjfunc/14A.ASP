<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'发招(比拼内力)♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
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
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【抢夺宝物】<font color=" & sayscolor & ">"+attack(mid(says,i+1),towhoid,towho)+"</font>"
call chatsay("抢夺宝物",towhoway,towho,saycolor,addwordcolor,addsays,says)

function attack(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 宝物 from 用户 where 姓名='" & sjjh_name &"'",conn
if rs("宝物")=Application("sjjh_baowuname") then
		Response.Write "<script language=JavaScript>{alert('你自己有宝物还要抢人家的呀！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close

'对方的判断
rs.open "select 宝物,保护,宝物修练 from 用户 where 姓名='" & to1 &"'",conn
 if rs("保护")=true then
 Response.Write "<script language=JavaScript>{alert('对方正在保护状态呀！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("宝物")=Application("sjjh_baowuname") then
		conn.execute "update 用户 set 宝物修练=0,宝物='"& Application("sjjh_baowuname") &"' where 姓名='" & sjjh_name &"'"
		conn.execute "update 用户 set 宝物修练=0,宝物='无' where 姓名='" & to1 &"'"
	attack=xinxi & "##把%%的宝物:"& Application("sjjh_baowuname") &"抢走，因得到此宝。江湖宝物需要进行修练才可以得到更多的东西！"&gjtx
		
else
attack=xinxi & "##想抢%%的宝物,可是%%身并没有什么宝物啊!"&gjtx
		
end if
rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
end function

%>