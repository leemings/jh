<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'传授武功
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_cs then
	Response.Write "<script language=JavaScript>{alert('提示：传武功最少需要["&jhdj_cs&"]级,才可以操作！');}</script>"
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
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【传授武功】<font color=" & saycolor & ">"+cuan(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'传授武功
function cuan(fn1,to1)
fn1=abs(fn1)
if fn1<200 then
	Response.Write "<script language=JavaScript>{alert('提示：传送武功一次最少200的，别太小气！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 等级,grade,武功,门派 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if fn1>50000 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：传武功不要大于50000好不！');}</script>"
	Response.End
end if
if rs("门派")="出家" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是出家人不能操作！');}</script>"
	Response.End
end if
if rs("门派")="官府" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是官府的人不能操作！');}</script>"
	Response.End
end if
if rs("武功")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的武功不足！');}</script>"
	Response.End
end if
rs.close
rs.open "select 武功 FROM 用户 WHERE 姓名='" & to1 &"'",conn
if rs("武功")>1500000 then
Response.Write "<script language=JavaScript>{alert('失败，他武功已经大于150万了，根本不需要你保护了!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 武功=武功-" & fn1 & " where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 武功=武功+" & fn1 & " where 姓名='" & to1 &"'"
cuan= "##发功，满脸通红，呼呼直喘,头上青烟直冒，终于把" & fn1 & "的武功传给了%%了，%%万分感谢！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>	
conn.close
set conn=nothing
end function
%>