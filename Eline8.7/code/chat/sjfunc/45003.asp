<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'寻水晶球♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_xunshuijing then
	Response.Write "<script language=JavaScript>{alert('寻水晶球，需要江湖等级["&jhdj_xunshuijing&"]级的才可以操作！');}</script>"
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
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【寻水晶球】<font color=" & saycolor & ">"+xunshuijing(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'寻水晶球
function xunshuijing(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 法力,操作时间 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<(int(rnd*3)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*3)+1)-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
fla=rs("法力")
if rs("法力")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的法力不够无法施展呀，至少也得1000点啊！');window.close();}</script>"
	response.end
end if
rs.close
conn.execute "update 用户 set 操作时间=now()  where 姓名='"&sjjh_name&"'"
randomize 
r=int(rnd*9)+1
select case r
case 1
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么水晶球," & sjjh_name & "损耗100点法力!" 
	conn.execute "update 用户 set 法力=法力-100 where 姓名='" & sjjh_name &"'"
case 2
xunshuijing=sjjh_name & "你千辛万苦寻访<font color=red>水晶球</font>的下落，却不料在一条水沟里被你找到了，赶紧洗洗水晶球挖掘它的魔力吧."
	rs.open "SELECT w9 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w9"),"水晶球",1)
	conn.execute "update 用户 set  w9='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close

case 3
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么水晶球," & sjjh_name & "损耗100点法力!" 
	conn.execute "update 用户 set 法力=法力-100 where 姓名='" & sjjh_name &"'"
	
case 4
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么水晶球," & sjjh_name & "损耗100点法力!" 
	conn.execute "update 用户 set 法力=法力-100 where 姓名='" & sjjh_name &"'"

case 5
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么水晶球," & sjjh_name & "损耗100点法力!" 
	conn.execute "update 用户 set 法力=法力-100 where 姓名='" & sjjh_name &"'"

case 6
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么水晶球," & sjjh_name & "损耗100点法力!" 
	conn.execute "update 用户 set 法力=法力-100 where 姓名='" & sjjh_name &"'"

case 7
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么水晶球," & sjjh_name & "损耗100点法力!" 
	conn.execute "update 用户 set 法力=法力-100 where 姓名='" & sjjh_name &"'"

case 8
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么水晶球," & sjjh_name & "损耗100点法力!" 
	conn.execute "update 用户 set 法力=法力-100 where 姓名='" & sjjh_name &"'"

case 9
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么水晶球," & sjjh_name & "损耗100点法力!" 
	conn.execute "update 用户 set 法力=法力-100 where 姓名='" & sjjh_name &"'"
	
case 10
xunshuijing=sjjh_name & "你千辛万苦寻访<font color=red>水晶球</font>的下落，却不料在一条水沟里被你找到了，赶紧洗洗水晶球挖掘它的魔力吧."
	rs.open "SELECT w9 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w9"),"水晶球",1)
	conn.execute "update 用户 set  w9='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
