<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'魔幻水晶♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_shqiu then
	Response.Write "<script language=JavaScript>{alert('魔幻水晶，需要江湖等级["&jhdj_shqiu&"]级的才可以操作！');}</script>"
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
says="<font color=red>【魔幻水晶】<font color=" & saycolor & ">"+shqiu(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'魔幻水晶
function shqiu(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w9,法力 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
fla=rs("法力")

if rs("法力")<3000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的法力不够无法施展呀，至少也得1000点啊！');window.close();}</script>"
	response.end
end if
duyao=rs("w9")
if isnull(duyao) or duyao=abate(rs("w9"),"水晶球",1) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你连一个水晶球都没有也!');</script>"
	response.end
end if


rs.close
conn.execute "update 用户 set 操作时间=now()  where 姓名='"&sjjh_name&"'"

randomize 
r=int(rnd*5)+1
select case r
case 1
shqiu=sjjh_name & "从<bgsound src=wav/Ye150.wav loop=1>口袋里拿出水晶球，对着" & to1 & "口中唸唸有词，只见水晶球奇幻的散发出光芒照射在<font color=red>" & to1 & "</font>的身上，" & to1 & "迷迷糊糊的睡着了." 
   rs.open "SELECT w9 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3
   duyao=abate(rs("w9"),"水晶球",1)
   conn.execute "update 用户 set  法力=法力-3000,w9='"&duyao&"' where 姓名='"&sjjh_name&"'"
   n=Year(date())
   y=Month(date())
   r=Day(date())
   s=Hour(time())
   f=Minute(time())
   m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
application("sjjh_dianxuename")=application("sjjh_dianxuename")& to1 & "|"&sj2&"|"&";"
rs.close	

case 2
shqiu=sjjh_name & "从<bgsound src=wav/Ye150.wav loop=1>口袋里拿出水晶球，对着" & to1 & "口中唸唸有词，只见水晶球奇幻的散发出光芒照射在<font color=red>" & to1 & "</font>的身上，但这次水晶球失效." & to1 & "觉得好高兴喔水晶球对我没作用 " 
		conn.execute "update 用户 set 法力=法力-1000 where 姓名='" & sjjh_name &"'"

case 3
shqiu=sjjh_name & "从<bgsound src=wav/Ye150.wav loop=1>口袋里拿出水晶球，对着" & to1 & "口中唸唸有词，只见水晶球奇幻的散发出光芒照射在<font color=red>" & to1 & "</font>的身上，但这次水晶球失效." & to1 & "觉得好高兴喔水晶球对我没作用 " 
	conn.execute "update 用户 set 法力=法力-1000 where 姓名='" & sjjh_name &"'"
	
case 4
shqiu=sjjh_name & "从<bgsound src=wav/Ye150.wav loop=1>口袋里拿出水晶球，对着" & to1 & "口中唸唸有词，只见水晶球奇幻的散发出光芒照射在<font color=red>" & to1 & "</font>的身上，但这次水晶球失效." & to1 & "觉得好高兴喔水晶球对我没作用 " 
	conn.execute "update 用户 set 法力=法力-1000 where 姓名='" & sjjh_name &"'"

case 5
shqiu=sjjh_name & "从<bgsound src=wav/Ye150.wav loop=1>口袋里拿出水晶球，对着" & to1 & "口中唸唸有词，只见水晶球奇幻的散发出光芒照射在<font color=red>" & to1 & "</font>的身上，但这次水晶球失效." & to1 & "觉得好高兴喔水晶球对我没作用 " 
	conn.execute "update 用户 set 法力=法力-1000 where 姓名='" & sjjh_name &"'"
	
case 6
shqiu=sjjh_name & "从<bgsound src=wav/Ye150.wav loop=1>口袋里拿出水晶球，对着" & to1 & "口中唸唸有词，只见水晶球奇幻的散发出光芒照射在<font color=red>" & to1 & "</font>的身上，" & to1 & "迷迷糊糊的睡着了." 
	 rs.open "SELECT w9 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3
   duyao=abate(rs("w9"),"水晶球",1)
   conn.execute "update 用户 set  法力=法力-5000,w9='"&duyao&"' where 姓名='"&sjjh_name&"'"
   conn.execute "update 用户 set 状态='眠',登录=now()+1,事件原因='"&sjjh_name&" 眠:"&fn1&"' where 姓名='" & to1 &"'"
   call boot(to1,"监禁，操作者："&sjjh_name&","&fn1)
   conn.execute "insert into l(b,c,d,a,e) values ('" & sjjh_name & "','" & to1 & "','眠',now(),'" & fn1 & "')"
   rs.close	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>


