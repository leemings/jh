<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'拜师♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_bs then
	Response.Write "<script language=JavaScript>{alert('提示：拜师需要["&jhdj_bs&"]级才可以操作！');}</script>"
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
says="<font color=green>【拜师】<font color=" & saycolor & ">"+bais(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'拜师
function bais(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 师傅,银两,等级 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
sf=rs("师傅")
if sjjh_name=to1 then
	if sf="" or sf="无" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：你没有师傅，无法脱离师门！');}</script>"
		Response.End
	end if
if rs("银两")<80000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你没有8万块，人家不干呀！！');}</script>"
	Response.End
end if
	conn.execute "update 用户 set 银两=银两-50000,师傅='无',保留1='保留' where 姓名='" & sjjh_name &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	bais="##向原师傅"& sf &"交纳了8万块的分手费，终于与"& sf &"脱离了师徒关系！"
	exit function
end if
if sf=to1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：["& to1&"]已经是你师傅了');}</script>"
	Response.End
end if
if rs("师傅")<>"" and rs("师傅")<>"无" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：想拜["& to1&"]为师，请与现在师傅脱离关系！');}</script>"
	Response.End
end if
if rs("银两")<50000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你没有5万两["& to1&"]不想收你为弟子！');}</script>"
	Response.End
end if
dj=rs("等级")
rs.close	
rs.open "select 等级 FROM 用户 WHERE 姓名='" & to1 &"'",conn,2,2
if rs("等级")<30 or dj>rs("等级") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：["& to1&"]还不够30级呀,不能收弟子！');}</script>"
	Response.End
end if
bais="##正准备向%%交纳了5万块拜师费，提出了拜师申请，也不知道人家同意不！"
Application.Lock
Application("sjjh_bais_sf")=to1
Application("sjjh_bais_td")=sjjh_name
Application.UnLock
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
