<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'蓝玛瑙♀wWw.happyjh.com♀
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
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【蓝玛瑙】<font color=" & saycolor & ">"+lbsw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function lbsw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w6,等级,体力,职业,法力 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<100 or fn1>200000 then
Response.Write "<script language=JavaScript>{alert('玛瑙一次最少100最多不能超过200000！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("职业")<>"魔法师" then
	Response.Write "<script language=JavaScript>{alert('你只是普通子民呢，不能使用蓝玛瑙！！请去职业转换为魔法师！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
duyao=rs("w6")
if isnull(duyao) or duyao=abate(rs("w6"),"蓝玛瑙",1)  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你连一棵蓝玛瑙也没有耶.看清楚好吗!');</script>"
	response.end
end if
if rs("等级")<80 then
	Response.Write "<script language=JavaScript>{alert('你的等级不够呀！使用宝物你还得练练呢要求80级以上！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1>200000  then
	Response.Write "<script language=JavaScript>{alert('别太贪了，最多别超过200000！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("法力")<2000 then
	Response.Write "<script language=JavaScript>{alert('你法力不够了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 体力,法力,w6 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn
if rs("体力")>5000000 then
lbsw=sjjh_name & " 你的体力够高的了，还要补吗？嘿嘿嘿.(你个人体力大于5000000)！"
else
duyao=abate(rs("w6"),"蓝玛瑙",1)
conn.execute "update 用户 set w6='"&duyao&"',体力=体力+"&fn1&",法力=法力-2000 where 姓名='" & sjjh_name &"'"
lbsw="传闻<font color=red size=2>" & sjjh_name & "</font> 使用了发光的<font color=blue>蓝玛瑙</font>体力狂增"&fn1&"点."
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>