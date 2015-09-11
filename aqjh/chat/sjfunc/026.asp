<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'杀气锐减
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
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
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=fakuan(mid(says,i+1)+0,towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'杀气锐减"
function fakuan(fn1,to1)
fn1=abs(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 师傅,智力,银两,杀人数,总杀人,道德,会员金卡 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("会员金卡")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[%%]对[<font color=red>##</font>]说：你哪里有<font color=red>"&fn1&"</font>块金卡？搞错了，你比我还神棍啊！"
	exit function
end if
if rs("道德")<fn1*1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[<font color=red>%%</font>]对[<font color=red>##</font>]说：你道德少于<font color=red>"&fn1*1000&"</font>，天界诸神都看不起你了，还怎么会帮你呢？"
	exit function
end if
rs.close
rs.open "SELECT 银两,grade,师傅,道德,总杀人,会员金卡,知质,智力,杀人数 FROM 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if rs("杀人数")>=3 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="<font color=red>[<b>官府</b>]对[##]<b>警告</b>说：你今天杀了<img src='picwords/3.gif'>个人了，不能减少杀气，你太坏了~！没把你抓起来算好了~！"
	exit function
end if
if rs("总杀人")<=0 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="<font color=" & saycolor & ">##欲使用[<b><font color=red>杀气锐减</font></b>]减少自己杀气，但##你的杀人数是<font color=red>0</font>了，你还要怎么减呀~你看[%%]在取笑你呢~~嘿嘿~~"
	exit function
end if
if fn1>20 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="<font color=" & saycolor & ">##欲使用[<b><font color=red>杀气锐减</font></b>]减少自己杀气，但爱情天界规定最大销减数为<font color=red>20</font>哦,你看[%%]在取笑你呢~~嘿嘿~~"
	exit function
end if
if rs("会员金卡")>=fn1 then
    conn.execute("update 用户 set 知质=知质+" & fn1 & ",道德=道德-" & fn1*1000 & ",总杀人=总杀人-" & fn1 & ",会员金卡=会员金卡-" & fn1 & " where 姓名='"&aqjh_name&"'")
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=red>【杀气锐减】</font>[<font color=blue>##</font>]从自己财产里拿出了[<font color=red>" & fn1 & "</font>]块金卡，买通了天界负责屠杀的天使恶魔，顺利的减少了[<font color=red>" & fn1 & "</font>]个杀人总数，知质增加[<font color=red>" & fn1 & "</font>],道德减少了[<font color=red>" & fn1*1000 & "</font>]真是天作孽呀~"
	exit function
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
