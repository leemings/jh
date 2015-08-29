<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'魔法特效♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_jhdj<20 then
	Response.Write "<script language=JavaScript>{alert('提示：想使用魔法特效需要[20]级才行！');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
call dianzan(towho)
act=0
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
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))

select case mycz
    case "/字体魔法"
        says="<font color=green>【字体魔法】</font><font color=" & saycolor & ">"+mofa1(mid(says,i+1),towho)+"</font>"
    case "/移动魔法"
	says="<font color=green>【移动魔法】</font><font color=" & saycolor & ">"+mofa2(mid(says,i+1),towho)+"</font>"
    case "/按钮魔法"
        says="<font color=green>【按钮魔法】</font><font color=" & saycolor & ">"+mofa3(mid(says,i+1),towho)+"</font>"
end select
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'字体魔法
function mofa1(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('提示：滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if
zt=split(fn1,"|")
if ubound(zt)<>3 then
	Response.Write "<script language=JavaScript>{alert('提示：格式为：内容|颜色|大小|字体 \n\n    如：我爱你|red|20|黑体 ');}</script>"
	Response.End 
end if

if not isnumeric(zt(2)) then 
	Response.Write "<script language=JavaScript>{alert('提示：操作错误，[大小]请使用数字！');}</script>"
	Response.End 
end if
fontgb=trim(zt(0))
fontcolor=trim(zt(1))
fontsize=abs(int(clng(zt(2))))
fontt=trim(zt(3))
fontx=len(fontgb)*50
if fontsize>30 and fontsize<0 then
	Response.Write "<script language=JavaScript>{alert('提示：字体大小不得小于0且大于30');}</script>"
	Response.End 
end if

Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select 体力,内力,法力 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("体力")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：字体魔法需要体力1000点！');}</script>"
	Response.End
end if
if rs("内力")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：字体魔法需要内力1000点！');}</script>"
	Response.End
end if

if rs("法力")<fontx then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：字体魔法需要法力"&fontx&"点！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 体力=体力-1000,内力=内力-1000,法力=法力-"&fontx&" where 姓名='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing

mofa1="##使出了字体魔法(共花费体力1000点,内力1000点,法力"&fontx&"点)：<br><font color="&fontcolor&" size="& fontsize&" face="&fontt&">"&fontgb&"</font>"
end function

'移动魔法
function mofa2(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('提示：滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if
zt=split(fn1,"|")
if ubound(zt)<>5 then
	Response.Write "<script language=JavaScript>{alert('提示：格式为：内容|颜色|背景色|left/right|速度|延时 \n\n    如：我爱你|red|blue|right|20|500 ');}</script>"
	Response.End 
end if

if not isnumeric(zt(4)) then 
	Response.Write "<script language=JavaScript>{alert('提示：操作错误，[速度]请使用数字！');}</script>"
	Response.End 
end if
if not isnumeric(zt(5)) then 
	Response.Write "<script language=JavaScript>{alert('提示：操作错误，[延时]请使用数字！');}</script>"
	Response.End 
end if

gdgb=trim(zt(0))
gdcolor=trim(zt(1))
gdbg=trim(zt(2))
gdtype=trim(zt(3))
gdspeed=abs(int(clng(zt(4))))
gdys=abs(int(clng(zt(5))))

if gdtype<>"left" and gdtype<>"right" then
	Response.Write "<script language=JavaScript>{alert('提示：滚动方式只有二种：left或right');}</script>"
	Response.End 
end if

if gdspeed<0 or gdspeed>50 then
	Response.Write "<script language=JavaScript>{alert('提示：速度大于0并小于50');}</script>"
	Response.End 
end if
if gdys<100 or gdys>500 then
	Response.Write "<script language=JavaScript>{alert('提示：延时大于100并小于500');}</script>"
	Response.End 
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select 体力,内力,法力 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("体力")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：移动魔法需要体力1000点！');}</script>"
	Response.End
end if
if rs("内力")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：移动魔法需要内力1000点！');}</script>"
	Response.End
end if

if rs("法力")<500 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：移动魔法需要法力500点！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 体力=体力-1000,内力=内力-1000,法力=法力-500 where 姓名='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing

mofa2="##使出了移动魔法(共花费体力1000点,内力1000点,法力500点)：<br><font color="&gdcolor&"><marquee bgcolor="&gdbg&" direction="&gdtype&" scrollamount="&gdspeed&" scrolldelay="&gdys&">"&gdgb&"</marquee></font>"

end function

'按钮魔法
function mofa3(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('提示：滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if
zt=split(fn1,"|")
if ubound(zt)<>2 then
	Response.Write "<script language=JavaScript>{alert('提示：格式为：内容|颜色|弹出内容 \n\n    如：我爱你|red|喜欢你 ');}</script>"
	Response.End 
end if


angb=trim(zt(0))
ancolor=trim(zt(1))
angb2=trim(zt(2))



Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select 体力,内力,法力 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("体力")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：按钮魔法需要体力1000点！');}</script>"
	Response.End
end if
if rs("内力")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：按钮魔法需要内力1000点！');}</script>"
	Response.End
end if

if rs("法力")<500 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：按钮魔法需要法力500点！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 体力=体力-1000,内力=内力-1000,法力=法力-500 where 姓名='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing

mofa3="##使出了按钮魔法(共花费体力1000点,内力1000点,法力500点)：<br><input type=button value="&angb&" style=background-color:"&ancolor&" onclick=window.alert('"&angb2&"');>"

end function
%>