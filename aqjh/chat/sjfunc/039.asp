<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'高级奖励
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
if towho=aqjh_name or towho=application("aqjh_automanname") then
	towho="大家"
else
	call dianzan(towho)
end if
if towho="大家" then
	Response.Write "<script language=JavaScript>{alert('提示：操作错误~别想作弊~！！！');}</script>"
	Response.End
end if
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
says="<font color=red>【玩家奖励】<font color=blue>"+lbsw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function lbsw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w5,等级,职业,法力,道德,转生 FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<1 or fn1>10 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>##</font>你发傻啊，随便乱按，找死！！！不要随便乱按东东~！！！</font>"
	exit function
end if
if rs("等级")<90 and rs("转生")<1 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>##</font>你等级太低了，没资格使用此高—级—功—能—啦~~`~！"
	exit function
end if
if fn1>10 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>##</font>你发傻啊，随便乱按，找死！！！不要随便乱按东东~！！！</font>"
	exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select allvalue,法力,道德,等级,w5,会员金卡,记录2 FROM 用户 WHERE 姓名='" & to1 &"'" ,conn
mdate=rs("记录2")
if rs("记录2")=true then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>%%</font>他已经有过奖励了，<font color=blue>##</font>你就别操心了，看看还有没人需要帮忙的~"
	exit function
end if
if rs("等级")<120 then
lbsw=to1 & "的等级太低了,不够120级，不能奖励！"
else
zstemp=add(rs("w5"),"养猪卡",5)
conn.execute "update 用户 set 记录2=true,w5='"&zstemp&"',会员金卡=会员金卡+120 where 姓名='" & to1 &"'"
lbsw="谢谢<font color=blue>[%%]</font>您的支持，您已经120以上，<font color=blue>[##]</font>我代表快乐谢谢你，并奖励您养猪卡<font color=red>5</font>张，快去养您的小猪猪吧，再送您会员金卡<font color=red>120</font>块，多带朋友来玩哦！</font>"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>%>