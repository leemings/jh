<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'魔法石
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
says="<font color=red>【魔法石】<font color=" & saycolor & ">"+lbsw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function lbsw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,等级,allvalue,职业,法力,道德,转生 FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<1 or fn1>10 then
Response.Write "<script language=JavaScript>{alert('魔法石一次最少1最多不能超过10个！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("职业")<>"魔法师" then
	Response.Write "<script language=JavaScript>{alert('你只是普通子民呢，不能使用魔法石！！请去职业转换为魔法师！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
duyao=rs("w6")
if isnull(duyao) or duyao=abate(rs("w6"),"魔法石",fn1*10)  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你连"&fn1*10&"粒魔法石也没有耶.看清楚好吗!');</script>"
	response.end
end if
if rs("等级")<30 then
	Response.Write "<script language=JavaScript>{alert('你的等级不够呀！使用宝物你还得练练呢要求30级以上！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("转生")>2 then
	Response.Write "<script language=JavaScript>{alert('你已经是多次转生人了，还不厉害啊，去照顾新人吧！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1>10  then
	Response.Write "<script language=JavaScript>{alert('别太贪了，最多别超过10个！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("法力")<fn1*150000 then
	Response.Write "<script language=JavaScript>{alert('你法力不够"&fn1*150000&"万！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("道德")<fn1*10000 then
	Response.Write "<script language=JavaScript>{alert('你道德不够"&fn1*10000&"！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select allvalue,法力,道德,等级,w6 FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn
if rs("等级")>200 then
lbsw=aqjh_name & " 你的总积分够高的了，还要补吗？嘿嘿嘿.(你个人等级大于200)！"
else
duyao=abate(rs("w6"),"魔法石",fn1*10)
conn.execute "update 用户 set w6='"&duyao&"',allvalue=allvalue+"&fn1*1000&",法力=法力-"&fn1*150000&",道德=道德-"&fn1*10000&" where 姓名='" & aqjh_name &"'"
lbsw="传闻<font color=red size=2>" & aqjh_name & "</font>使用了传说中的妖界镇妖之宝—<font color=blue>魔法石</font>，总积分狂增<font color=red>"&fn1*1000&"</font>点，道德下降<font color=red>"&fn1*10000&"</font>点,法力减少<font color=red>"&fn1*150000&"</font>点，看得妖界众多小鬼心里直打颤，真是阴损的恐怖魔法………"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>%>