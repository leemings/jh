<%@ LANGUAGE=VBScript codepage ="936" %>

<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'讨取轻功
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以偷钱！');}</script>"
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
says=Replace(says,"&amp;","&")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【讨取轻功】<font color=" & saycolor & ">"+demand(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'讨取轻功
function demand(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 保护,等级,姓名,会员等级,轻功,银两,grade from 用户 where 姓名='" & sjjh_name &"'" ,conn,2,2
if rs("保护")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你正在练功保护请不要使用讨取轻功!');}</script>"
	Response.End
end if
fla=rs("轻功")
if fla<500 then
Response.Write "<script language=JavaScript>{alert('你的轻功不够无法施展呀，至少也得500点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

rs.close
rs.open "select 保护,等级,姓名,会员等级,银两,grade from 用户 where 姓名='" & to1 &"'" ,conn,2,2
if rs("保护")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方正在练功保护请不要对他使用讨取!');}</script>"
	Response.End
end if
if rs("等级")<=10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('初入江湖不能对他讨取');}</script>"
	Response.End
end if
if rs("grade")>=10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你要对官府讨取吗我看不要想太多了！？！');}</script>"
	Response.End
end if
rs.close
rs.open "select 性别,轻功 FROM 用户 WHERE 姓名='" & to1 &"'",conn
xb=rs("性别")
yin=int(rs("轻功")/3)
if xb="男" then
Response.Write "<script language=JavaScript>{alert('此功能不能用在男的身上！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 轻功=轻功-"&yin&" where 姓名='" & to1 &"'"
conn.execute "update 用户 set 轻功=轻功+"&yin&" WHERE  姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 轻功=轻功-500 WHERE  姓名='" & sjjh_name &"'"
demand=sjjh_name & "<bgsound src=wav/Phant006.wav loop=1>贼头贼脑的对" & to1 & "说：这位侠客你好你的轻功好像太多了可否暂借一点轻功？急用，求你了...说得<font color=red>" & to1 & "</font>不好意思地从身上拿出三分之一的轻功给了" & sjjh_name & "说,本人轻功多的要命，拿点去吧.不用还了. " & sjjh_name & "因此得到"&yin&"点轻功，轻功消耗500点"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>
