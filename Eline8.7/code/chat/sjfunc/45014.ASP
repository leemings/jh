<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%'血滴子♀wWw.happyjh.com♀
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
says="<font color=green>【血滴子】<font color=" & sayscolor & ">"+xuedizi(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function xuedizi(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w6,w9,门派,法力,等级 FROM 用户 WHERE  姓名="&"'" & sjjh_name &"'",conn,3,3
w6w=rs("w9")
if rs("门派")="官府"  then
Response.Write "<script language=JavaScript>{alert('你是官府人员啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("等级")<20 then
Response.Write "<script language=JavaScript>{alert('此功能需要20级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("法力")<8000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得8000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 门派 FROM 用户 WHERE 姓名="&"'" & towho &"'",conn,3,3
if rs("门派")="官府"  then
Response.Write "<script language=JavaScript>{alert('失败，你不能对高级管理员或站长特封的人员操作!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
wuyn=iswp(w6w,"血滴子")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('你有血滴子吗？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
fq=abate(w6w,"血滴子",1)
conn.execute "update 用户 set  w9='"&fq&"' where 姓名='"&sjjh_name&"'"

conn.execute "update 用户 set 法力=法力-8000 where  姓名="&"'" & sjjh_name &"'"

conn.execute "update 用户 set 体力=体力-2000 where 姓名="&"'" & towho &"'"
conn.execute "update 用户 set 内力=内力-2000 where  姓名="&"'" & towho &"'"
rs.open "select * from 用户 where 姓名="&"'" & towho &"'",conn
if rs("体力")<0  then
conn.execute "update 用户 set 杀人数=杀人数+1 where 姓名="&"'" & sjjh_name &"'"
conn.execute "update 用户 set 状态='死' where 姓名="&"'" & towho &"'"
xuedizi="##抛出法器血滴子，<bgsound src=wav/Bombs020.wav loop=1>血滴子罩在了<font color=red>%%</font>头上，%%被血滴子击中法力全无，体力消耗2000，内力失去2000，死翘翘了..." 
 call boot(towho,"血滴子，操作者："&sjjh_name&","&fn1)
else
xuedizi="##抛出法器<bgsound src=wav/Bombs020.wav loop=1>血滴子，血滴子罩在了<font color=red>%%</font>头上，%%被血滴子击中法力全无，体力消耗2000，内力失去2000，没死已经算命大了..." 
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>