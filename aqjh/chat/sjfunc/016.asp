<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'卡片合成
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
says="<font color=red>【卡片合成】<font color=" & saycolor & ">"+peibashi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'卡片合成
function peibashi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 智力,法力,等级,w5,职业,金币 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
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
if rs("智力")<100 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的智力不够无法施展呀，至少也得100点啊！');window.close();}</script>"
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
if dj<80 then
Response.Write "<script language=JavaScript>{alert('此功能需要[80]级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

duyao=rs("w5")
if isnull(duyao) or duyao=abate(rs("w5"),"涨钱卡",1) or duyao=abate(rs("w5"),"练功卡",1) or duyao=abate(rs("w5"),"宝物卡",1) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：三种卡片缺一不可看清楚在来!');</script>"
	response.end
end if

rs.close
rs.open "select w5 from 用户 where 姓名='"&aqjh_name&"'",conn
duyao=abate(rs("w5"),"涨钱卡",1)
conn.execute "update 用户 set  w5='"&duyao&"' where 姓名='"&aqjh_name&"'"
duyao=abate(rs("w5"),"练功卡",1)
conn.execute "update 用户 set  w5='"&duyao&"' where 姓名='"&aqjh_name&"'"
duyao=abate(rs("w5"),"宝物卡",1)
conn.execute "update 用户 set  w5='"&duyao&"' where 姓名='"&aqjh_name&"'"
duyao=add(rs("w5"),"升级卡",1)
conn.execute "update 用户 set 智力=智力-100,金币=金币-2,w5='"&duyao&"' where 姓名='"&aqjh_name&"'"
peibashi=aqjh_name & "<bgsound src=wav/m2.mid loop=1>取出涨钱卡、练功卡、宝物卡三种无敌卡片，三种卡片纠集在一起，一道黑色光芒升起，三种卡片唰的一声，化做一到白烟消失得无影无踪，只见地上遗留了当年卡片之王发明的方术族族宝<font color=red>升级卡<img src=card/ywk.gif></font>1张," & aqjh_name & "机缘巧合之下获得的至宝有可能使他增加自身强大的能量，狂傲快乐....."
 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>%>