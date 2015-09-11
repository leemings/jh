<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%
'用钱砸人
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=session("aqjh_name")
aqjh_grade=session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
%>
<!--#include file="pkif.asp"-->
<%
if aqjh_jhdj<20 then
	Response.Write "<script language=JavaScript>{alert('提示:需要20等级！');}</script>"
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
'对暂离开、点哑穴判断
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【拿钱砸人】<font color=red>"+za(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function za(fn1,to1)
fn1=int(abs(fn1))
if to1=aqjh_name or to1="大家" then
   Response.Write "<script language=JavaScript>{alert('提示:不能对自己和大家操作？');}</script>"
   response.end
end if
if fn1<=1000 then
   Response.Write "<script language=JavaScript>{alert('提示:你也太小气了吧,没有1000两以上是拿不出手的!');}</script>"
   response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
mygrade=rs("grade")
if rs("银两")<fn1 then
	Response.Write "<script language=JavaScript>{alert('提示:你哪里有"&fn1&"两银子？');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("门派")="官府" and mygrade<10 then
	Response.Write "<script language=JavaScript>{alert('提示:你是官府的,就不要再捣乱了!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("门派")="出家"  then
Response.Write "<script language=JavaScript>{alert('失败，你是出家人想要做什么，找死啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select * from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("等级")<30 or rs("体力")<fn1*0.05 then
	Response.Write "<script language=JavaScript>{alert('提示:人家等级和体力那么低,再给你砸一下就死翘翘了!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("门派")="官府" and mygrade<10  then
	Response.Write "<script language=JavaScript>{alert('提示:不能攻击官府人员!小心被抓!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("门派")="出家"  then
Response.Write "<script language=JavaScript>{alert('失败，你想对出家人做什么，找死啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
conn.execute "update 用户 set 道德=道德-20,银两=银两-"&fn1&" where 姓名='" & aqjh_name &"'"
randomize timer
s=int(rnd*2+1)
tl=int(fn1*0.05)
if s=1 then
za="##<font color=#ff0000>拿出"&fn1&"银两嗖的一声向<font color=blue>"&to1&"</font>(体力减"&tl&")扔了过去<img height=50 src=img/251.gif>，砸得"&to1&"头破血流，哈哈,打不过你用钱砸死你~~~~！！由于得意忘形，自己的道德下降了20，有得必有失嘛！</font>"
 	conn.execute "update 用户 set 体力=体力-"&tl&" where 姓名='" & to1 &"'"
else
za="##<font color=#ff0000>拿出"&fn1&"银两嗖的一声向<font color=blue>"&to1&"</font>扔了过去<img height=50 src=img/251.gif>，手法太烂没砸到,还下降了20道德，亏大了~~~~！！没想到<font color=blue>"&to1&"</font>来了一个鲤鱼跳龙门，双手趴地，正好按住了<font color=blue>"&fn1&"</font>两银子！哇，天上掉馅饼啊！从他的笑当中，在场的朋友顿时领悟到了花儿为什么会这样红！</font>"
	conn.execute "update 用户 set 银两=银两+"&fn1&" where 姓名='" & to1 &"'"
end if
set rs=nothing	
conn.close
set conn=nothing
end function
%>
