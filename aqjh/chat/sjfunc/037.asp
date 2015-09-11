<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'寻找金币
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_jhdj<jhdj_xunshuijing then
	Response.Write "<script language=JavaScript>{alert('寻找金币，需要江湖等级["&jhdj_xunshuijing&"]级的才可以操作！');}</script>"
	Response.End
end if
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
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【寻找金币】<font color=" & saycolor & ">"+xunshuijing(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'寻找金币
function xunshuijing(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 法力,等级,操作时间,知质,职业 FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn,2,2
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
if rs("法力")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的法力不够无法施展呀，至少也得5000点啊！');window.close();}</script>"
	response.end
end if
if rs("知质")<20 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的知质不够无法施展呀，至少也得20点啊！');window.close();}</script>"
	response.end
end if
if rs("职业")<>"冒险家" then
	Response.Write "<script language=JavaScript>{alert('你非冒险家，不能找宝物！！请去职业转换为冒险家！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
dj=rs("等级")
if dj<30 then
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：等级为30级才可以进行寻找金币！');}</script>"
	Response.End
end if
rs.close
conn.execute "update 用户 set 操作时间=now(),法力=法力-5000,知质=知质-20 where 姓名='"&aqjh_name&"'"
randomize 
r=int(rnd*6)+1
select case r
case 1
xunshuijing=aqjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么金币," & aqjh_name & "损耗1000点法力!" 
	conn.execute "update 用户 set 法力=法力-1000 where 姓名='" & aqjh_name &"'"
	
case 2
xunshuijing=aqjh_name & "真幸运，你寻遍了江湖各个角落找到了寻找人100点法力" & aqjh_name & "增加了法力100点!" 
	conn.execute "update 用户 set 法力=法力+100 where 姓名='" & aqjh_name &"'"

case 3
xunshuijing=aqjh_name & "你寻遍了江湖各个角落终于找到了[1]个金币," & aqjh_name & "损耗1000点法力!" 
	conn.execute "update 用户 set 法力=法力-1000,金币=金币+1 where 姓名='" & aqjh_name &"'"

case 4
xunshuijing=aqjh_name & "真幸运，你寻遍了江湖各个角落找到了寻金币者100点法力" & aqjh_name & "增加了法力100点!" 
	conn.execute "update 用户 set 法力=法力+100 where 姓名='" & aqjh_name &"'"

case 5
xunshuijing=aqjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么金币," & aqjh_name & "损耗1000点法力!" 
	conn.execute "update 用户 set 法力=法力-1000 where 姓名='" & aqjh_name &"'"

case 6
xunshuijing=aqjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么金币," & aqjh_name & "损耗1000点法力!" 
	conn.execute "update 用户 set 法力=法力-1000 where 姓名='" & aqjh_name &"'"
	rs.close
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>