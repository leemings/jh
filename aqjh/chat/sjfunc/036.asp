<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'介绍奖励
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if instr(Application("aqjh_user"),aqjh_name)=0 then
	Response.Write "<script language=JavaScript>{alert('提示：你不是正站长不能操作！');}</script>"
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
says="<font color=red>【拉人奖励】<font color=blue>"+lbsw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function lbsw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w5,等级,职业,法力,道德,转生,门派 FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<1 or fn1>10 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>##</font>你别被他骗了，一次有<font color=red>" & fn1 & "</font>人过关吗？慢慢来奖励<font color=blue>[%%]</font>吧！~仓库只有<font color=red>10</font>的物资了！</font>"
	exit function
end if
if rs("门派")<>"官府" then
	Response.Write "<script language=JavaScript>{alert('拉人奖励只有官府的才可以操作，你想做官府么！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("等级")<90 and rs("转生")<1 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>##</font>虽然你是官府，但你等级也太低了吧~`等90级吧！"
	exit function
end if
if fn1>10 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>##</font>你别被他骗了，一次有<font color=red>" & fn1 & "</font>人过关吗？慢慢来奖励<font color=blue>[%%]</font>吧！~仓库只有<font color=red>10</font>的物资了！</font>"
	exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select allvalue,法力,道德,等级,w5 FROM 用户 WHERE 姓名='" & to1 &"'" ,conn
if rs("等级")<30 then
lbsw=to1 & "的等级太低了，不能被使用此功能奖励！"
else
zstemp=add(rs("w5"),"加点卡",fn1)
conn.execute "update 用户 set w5='"&zstemp&"',金币=金币+"&fn1*50&" where 姓名='" & to1 &"'"
lbsw="<font color=blue>[%%]</font>介绍玩家<font color=red>" & fn1 & "</font>位到爱情做客，经过<font color=blue>[##]</font>核查确认后，颁发奖励加点卡<font color=red>" & fn1 & "</font>张，金币<font color=red>" &fn1*50& "</font>个，谢谢支持爱情发展，多带朋友来玩哦！</font>"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>%>