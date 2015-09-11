<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'暗器合成
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
says="<font color=red>【暗器合成】<font color=" & saycolor & ">"+peibashi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'暗器合成
function peibashi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 法力,等级,w4,职业,金币 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
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
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的金币不够无法施展呀，至少也得2个啊！');window.close();}</script>"
	response.end
end if
if dj<60 then
Response.Write "<script language=JavaScript>{alert('此功能需要[60]级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

duyao=rs("w4")
if isnull(duyao) or duyao=abate(rs("w4"),"梅花镖",5) or duyao=abate(rs("w4"),"毒龙镖",5) or duyao=abate(rs("w4"),"吸血神镖",5) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：三个令牌缺一不可看清楚在来!');</script>"
	response.end
end if

rs.close
rs.open "select w4 from 用户 where 姓名='"&aqjh_name&"'",conn
duyao=abate(rs("w4"),"梅花镖",5)
conn.execute "update 用户 set  w4='"&duyao&"' where 姓名='"&aqjh_name&"'"
duyao=abate(rs("w4"),"毒龙镖",5)
conn.execute "update 用户 set  w4='"&duyao&"' where 姓名='"&aqjh_name&"'"
duyao=abate(rs("w4"),"吸血神镖",5)
conn.execute "update 用户 set  w4='"&duyao&"' where 姓名='"&aqjh_name&"'"
duyao=add(rs("w4"),"暴雨犁花针",2)
conn.execute "update 用户 set 金币=金币-2,w4='"&duyao&"' where 姓名='"&aqjh_name&"'"
peibashi=aqjh_name & "<bgsound src=wav/m2.mid loop=1>方术师取出梅花镖、毒龙镖、吸血神镖三种超强暗器，三种暗器纠集在一起，一道红色光芒升起，三种暗器砰的一声，化做一到白烟消失得无影无踪，只见地上遗留了当年千王之王发明的暗器之王<font color=red>暴雨犁花针</font>2只," & aqjh_name & "机缘巧合之下获得的至宝有可能使他狂傲江湖，真是好运随身来....."
 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>%>