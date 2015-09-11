<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'炼魔法石
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
says="<font color=red>【炼魔法石】<font color=" & saycolor & ">"+peibashi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'炼魔法石
function peibashi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 法力,等级,w6,职业,金币 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
dj=rs("等级")
fla=rs("法力")
if rs("法力")<75000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的法力不够无法施展呀，至少也得75000点啊！');window.close();}</script>"
	response.end
end if
if rs("职业")<>"炼金师" then
	Response.Write "<script language=JavaScript>{alert('你只是普通子民呢！！请去职业转换为炼金师！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("金币")<2 then
Response.Write "<script language=JavaScript>{alert('此功能需要2个金币呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj<30 then
Response.Write "<script language=JavaScript>{alert('此功能需要[30]级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

duyao=rs("w6")
if isnull(duyao) or duyao=abate(rs("w6"),"帅哥令",2) or duyao=abate(rs("w6"),"美人令",2) or duyao=abate(rs("w6"),"多情环",2) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：三个令牌缺一不可看清楚在来!');</script>"
	response.end
end if

rs.close
rs.open "select w6 from 用户 where 姓名='"&aqjh_name&"'",conn
duyao=abate(rs("w6"),"帅哥令",2)
conn.execute "update 用户 set  w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
duyao=abate(rs("w6"),"美人令",2)
conn.execute "update 用户 set  w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
duyao=abate(rs("w6"),"多情环",2)
conn.execute "update 用户 set  w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
duyao=add(rs("w6"),"魔法石",1)
conn.execute "update 用户 set 金币=金币-2,w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
peibashi=aqjh_name & "<bgsound src=wav/m2.mid loop=1>取出帅哥令、美人令、多情环三种宝物，三种宝物结合在一起，一道光芒<img src=img/hs1.gif>升起，三种宝物将会回归天界回复天命，" & aqjh_name & "机缘巧合之下获得天人天魔至尊遗留下的-<font color=red>魔法石</font>-，真是踩到狗屎运了....."
 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>%>