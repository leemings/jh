<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'寻找配药♀一线网络→wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
f=Minute(time())
if f<45 or f>59 then
	Response.Write "<script language=JavaScript>{alert('紫色彩虹现在还没有配好药，寻找配药时间为每小时的45-59分钟！现在为动武时间！');window.close();}</script>"
	Response.End 
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【寻找配药】<font color=" & saycolor & ">"+xunfaqi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'寻找配药
function xunfaqi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 等级,会员等级,金币 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
dj=rs("等级")

fla=rs("金币")
if rs("金币")<20 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的金币不够无法施展呀，至少也得20个啊！');window.close();}</script>"
	response.end
end if
if dj<200 then
Response.Write "<script language=JavaScript>{alert('此功能需要200级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
hydj=rs("会员等级")
if hydj<4 then
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：会员等级为4级才可以进行寻找配药！');}</script>"
	Response.End
end if

rs.close
conn.execute "update 用户 set 金币=金币-20,操作时间=now()  where 姓名='"&sjjh_name&"'"
randomize 
r=int(rnd*9)+1
select case r
case 1
xunfaqi=sjjh_name & "你在" & to1 & "江湖里找到站长<font color=#0000FF>紫色彩虹</font>丢掉的[<b><font color=red>大力丸</font></b>]一个，吃下可增加内力<font color=red>50</font>万。"
	rs.open "SELECT w8 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w8"),"大力丸",1)
	conn.execute "update 用户 set  w8='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close


case 2
xunfaqi=sjjh_name & "你在" & to1 & "江湖里找到站长<font color=#0000FF>紫色彩虹</font>丢掉的[<b><font color=red>归命丸</font></b>]一个，吃下可增加体力<font color=red>180</font>万。"
	rs.open "SELECT w8 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w8"),"归命丸",1)
	conn.execute "update 用户 set  w8='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close

	
case 3
xunfaqi=sjjh_name & "在江湖里找到站长<font color=#0000FF>紫色彩虹</font>丢掉的[<b><font color=red>大宝日霜</font></b>]一个，吃下可增加武功<font color=red>100</font>万。"
	rs.open "SELECT w8 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w8"),"大宝日霜",1)
	conn.execute "update 用户 set  w8='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
	
case 4
xunfaqi=sjjh_name & "在江湖里找到站长<font color=#0000FF>紫色彩虹</font>丢掉的[<b><font color=red>炸药</font></b>]一个，可以攻击对方体力减少<font color=red>一</font>半。"
	rs.open "SELECT w8 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w8"),"炸药",1)
	conn.execute "update 用户 set  w8='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
	
case 5
xunfaqi=sjjh_name & "在江湖里找到站长<font color=#0000FF>紫色彩虹</font>丢掉的[<b><font color=red>化尸水</font></b>]一个，可以攻击对方内力减少<font color=red>一</font>半。"
	rs.open "SELECT w8 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w8"),"化尸水",1)
	conn.execute "update 用户 set  w8='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close

case 6
xunfaqi=sjjh_name & "在江湖里找到站长<font color=#0000FF>紫色彩虹</font>丢掉的[<b><font color=red>蛊虫</font></b>]一个，吸取对方内力<font color=red>1/2</font>。"
	rs.open "SELECT w8 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w8"),"蛊虫",1)
	conn.execute "update 用户 set  w8='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close

case 7
xunfaqi=sjjh_name & "在江湖里找到站长<font color=#0000FF>紫色彩虹</font>丢掉的[<b><font color=red>化功散</font></b>]一个，吸取对方武功<font color=red>1/2</font>。"
	rs.open "SELECT w8 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w8"),"化功散",1)
	conn.execute "update 用户 set  w8='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close

case 8
xunfaqi=sjjh_name & "在江湖里找到站长<font color=#0000FF>紫色彩虹</font>丢掉的[<b><font color=red>吸血虫</font></b>]一个，吸取对方武功<font color=red>1/2</font>。"
	rs.open "SELECT w8 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w8"),"吸血虫",1)
	conn.execute "update 用户 set  w8='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
	
case 9
xunfaqi=sjjh_name & "在江湖里找到站长<font color=#0000FF>紫色彩虹</font>丢掉的[<b><font color=red>舍利子</font></b>]一个，使用后杀人数清<font color=red>0</font>。"
	rs.open "SELECT w8 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w8"),"舍利子",1)
	conn.execute "update 用户 set  w8='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
	
case 10
xunfaqi=sjjh_name & "在江湖里找到站长<font color=#0000FF>紫色彩虹</font>丢掉的[<b><font color=red>三花聚顶</font></b>]一个，使用后武功增加<font color=red>50</font>万。"
	rs.open "SELECT w8 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w8"),"三花聚顶",1)
	conn.execute "update 用户 set  w8='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
