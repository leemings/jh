<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'摇钱术♀wWw.happyjh.com♀
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
says="<font color=red>【摇钱术】<font color=" & saycolor & ">"+ycs(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function ycs(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 法力,w6,等级,体力,职业 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<100 or fn1>3000000 then
Response.Write "<script language=JavaScript>{alert('摇一次最少100最多不能超过3000000！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("职业")<>"魔法师" then
	Response.Write "<script language=JavaScript>{alert('你只是普通子民呢，不能使用摇钱术！！请去职业转换为魔法师！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
duyao=rs("w6")
if isnull(duyao) or duyao=abate(rs("w6"),"摇钱树",1)  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你连一棵摇钱树也没有耶.看清楚好吗!');</script>"
	response.end
end if

if rs("等级")<80 then
	Response.Write "<script language=JavaScript>{alert('你的等级不够呀！使用摇钱树你还得练练呢要求80级以上！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

if fn1>3000000 then
	Response.Write "<script language=JavaScript>{alert('别太贪了，最多别超过3百万！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("法力")<1000 then
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
rs.open "select w6,存款,法力 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn
if rs("存款")>400000000 then
ycs=sjjh_name & " 你还真贪耶！你这人真是贪心呀！...嘿嘿嘿.(你个人银两大于四亿)！"
else
duyao=abate(rs("w6"),"摇钱树",1)
conn.execute "update 用户 set w6='"&duyao&"',法力=法力-1000,存款=存款+"&fn1&" where 姓名='" & sjjh_name &"'"
ycs="恭喜恭喜,<font color=red size=2>" & sjjh_name & "</font>运气真好，不知道从那里找到了棵<font color=red size=2>摇钱树</font> ，只一摇就得了"&fn1&"两银子。"

end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

%>