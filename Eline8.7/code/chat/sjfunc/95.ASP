<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'寻找银币♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_xunshuijing then
	Response.Write "<script language=JavaScript>{alert('寻找银币，需要江湖等级["&jhdj_xunshuijing&"]级的才可以操作！');}</script>"
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
says="<font color=red>【寻找银币】<font color=" & saycolor & ">"+xunshuijing(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'寻找银币
function xunshuijing(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 法力,等级,操作时间 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<(int(rnd*5)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*5)+1)-sj
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
dj=rs("等级")
if dj<30 then
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：等级为30级才可以进行寻找银币！');}</script>"
	Response.End
end if
rs.close
conn.execute "update 用户 set 操作时间=now()  where 姓名='"&sjjh_name&"'"

randomize 
r=int(rnd*6)+1
select case r
case 1
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么银币," & sjjh_name & "损耗100点法力!" 
	conn.execute "update 用户 set 法力=法力-100 where 姓名='" & sjjh_name &"'"
	
case 2
xunshuijing=sjjh_name & "真幸运，你寻遍了江湖各个角落找到了寻找人10点法力" & sjjh_name & "增加了法力10点!" 
	conn.execute "update 用户 set 法力=法力+10 where 姓名='" & sjjh_name &"'"

case 3
xunshuijing=sjjh_name & "你寻遍了江湖各个角落终于找到了[1]个银币," & sjjh_name & "损耗100点法力!" 
	conn.execute "update 用户 set 法力=法力-100,银币=银币+1 where 姓名='" & sjjh_name &"'"

case 4
xunshuijing=sjjh_name & "真幸运，你寻遍了江湖各个角落找到了寻银币者10点法力" & sjjh_name & "增加了法力10点!" 
	conn.execute "update 用户 set 法力=法力+10 where 姓名='" & sjjh_name &"'"

case 5
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么银币," & sjjh_name & "损耗100点法力!" 
	conn.execute "update 用户 set 法力=法力-100 where 姓名='" & sjjh_name &"'"

case 6
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么银币," & sjjh_name & "损耗100点法力!" 
	conn.execute "update 用户 set 法力=法力-100 where 姓名='" & sjjh_name &"'"
	rs.close
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
