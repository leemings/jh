<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'配制令牌♀wWw.happyjh.com♀
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
says="<font color=red>【配制令牌】<font color=" & saycolor & ">"+peibashi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'配制令牌
function peibashi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 法力,等级,w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
dj=rs("等级")
fla=rs("法力")
if rs("法力")<20000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的法力不够无法施展呀，至少也得20000点啊！');window.close();}</script>"
	response.end
end if
if dj<50 then
Response.Write "<script language=JavaScript>{alert('此功能需要50级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

duyao=rs("w10")
if isnull(duyao) or duyao=abate(rs("w10"),"红钻石",1) or duyao=abate(rs("w10"),"绿钻石",1) or duyao=abate(rs("w10"),"蓝钻石",1) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：三个钻石缺一不可看清楚再来!');</script>"
	response.end
end if

rs.close
rs.open "select w10 from 用户 where 姓名='"&sjjh_name&"'",conn
duyao=abate(rs("w10"),"红钻石",1)
conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
duyao=abate(rs("w10"),"绿钻石",1)
conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
duyao=abate(rs("w10"),"蓝钻石",1)
conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
duyao=add(rs("w10"),"圣火令",1)
conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
peibashi=sjjh_name & "<bgsound src=wav/m2.mid loop=1>取出红、绿、蓝三种钻石，三种钻石结合在一起，一道光芒升起，<img src='img/look52.gif'>三种宝石将会化成武林圣火令<img src='img/e5.jpg' width='150' height='80'><br></font>" & sjjh_name & "这武林圣火令--他可号令天下.里面收藏了武林绝学有乾坤大挪移..七伤拳..九阳神功..寒冰绵掌等等..."
 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>