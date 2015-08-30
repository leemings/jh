<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'寻找金卡♀一线网络→wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_xunshuijing then
	Response.Write "<script language=JavaScript>{alert('寻找金卡，需要江湖等级["&jhdj_xunshuijing&"]级的才可以操作！');}</script>"
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
says=Replace(says,"&amp;","&")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【寻找金卡】<font color=" & saycolor & ">"+xunshuijing(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'寻找金卡
function xunshuijing(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 金币,会员等级 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
fla=rs("金币")
if rs("金币")<20 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的金币不够无法施展呀，至少也得20个啊！');window.close();}</script>"
	response.end
end if
hydj=rs("会员等级")
if hydj<5 then
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：会员等级为5级才可以进行寻找金卡！购买5级会员赠送400级（800万存点）每天可以领到100个金币！购买加QQ3494955 或打电话：0317-5465537');}</script>"
	Response.End
end if
rs.close
conn.execute "update 用户 set 操作时间=now()  where 姓名='"&sjjh_name&"'"
randomize 
r=int(rnd*6)+1
select case r
case 1
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么金卡," & sjjh_name & "损耗[<b><font color=red>20</font></b>]个金币!" 
	conn.execute "update 用户 set 金币=金币-20 where 姓名='" & sjjh_name &"'"

case 2
xunshuijing=sjjh_name & "你寻遍了江湖各个角落终于找到了[<b><font color=red>2</font></b>]张金卡," & sjjh_name & "真是幸运呀!" 
	conn.execute "update 用户 set 金卡=金卡+2 where 姓名='" & sjjh_name &"'"
	rs.close
case 3
xunshuijing=sjjh_name & "真幸运，你寻遍了江湖各个角落找到了寻找人的[<b><font color=red>5</font></b>]个金币" & sjjh_name & "增加了金币[<b><font color=red>5</font></b>]个!" 
	conn.execute "update 用户 set 金币=金币+5 where 姓名='" & sjjh_name &"'"

case 4
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么金卡" & sjjh_name & "损耗金币[<b><font color=red>5</font></b>]个!" 
	conn.execute "update 用户 set 金币=金币-5 where 姓名='" & sjjh_name &"'"

case 5
xunshuijing=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么金卡," & sjjh_name & "损耗[<b><font color=red>10</font></b>]个金币!" 
	conn.execute "update 用户 set 金币=金币-10 where 姓名='" & sjjh_name &"'"

case 6
xunshuijing=sjjh_name & "你寻遍了江湖各个角落终于找到了[<b><font color=red>5</font></b>]张金卡," & sjjh_name & "损耗[<b><font color=red>20</font></b>]个金币!" 
	conn.execute "update 用户 set 金币=金币-20,金卡=金卡+5 where 姓名='" & sjjh_name &"'"
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
