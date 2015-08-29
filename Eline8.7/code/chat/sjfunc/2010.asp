<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'帅哥令♀wWw.51eline.com♀
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
says="<font color=red>【帅哥令】<font color=" & saycolor & ">"+sgl(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function sgl(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 性别,w6,等级,体力,职业,法力 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
if rs("职业")<>"魔法师" then
	Response.Write "<script language=JavaScript>{alert('你只是普通子民呢，不能使用摇钱术！！请去职业转换为魔法师！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
duyao=rs("w6")
if isnull(duyao) or duyao=abate(rs("w6"),"帅哥令",1)  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你没有帅哥令耶.看清楚好吗!');</script>"
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
if rs("法力")<3000 then
	Response.Write "<script language=JavaScript>{alert('你法力不够3000点了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w6,法力,性别 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn
if rs("性别")="男" then
sgl=sjjh_name & " 你本来就是男同志，所以不能使用帅哥令阿,要不要给你买个眼镜呢！"
else
duyao=abate(rs("w6"),"帅哥令",1)
conn.execute "update 用户 set 性别='男',w6='"&duyao&"',法力=法力-3000 where 姓名='" & sjjh_name &"'"
sgl="<font color=red size=2>" & sjjh_name & " 突然亮出了 <font color=blue>帅哥令</font>!" & sjjh_name & " 在不知不觉中变成了帅哥哥....</font>"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

%>