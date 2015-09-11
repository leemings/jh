<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'申领奖励
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_name=Session("aqjh_name")
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
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【新人奖励】<font color=" & saycolor & ">"+sljk()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function sljk()
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if Session("sljk_time")="" then Session("sljk_time")=rs("lasttime")
mycd=DateDiff("n",Session("sljk_time"),now())
if mycd<60 then
    if DateDiff("n",rs("lasttime"),now())>=30 then
       sljk="你在一个小时前已经领过金卡了！"
    else
       sljk="你还没有泡够一个小时呢！"
    end if
    Response.Write "<script language=JavaScript>{alert('"&sljk&"');}</script>"
    Response.End
else
    conn.execute "update 用户 set 会员金卡=会员金卡+2,mvalue=mvalue-2000 where 姓名='" & aqjh_name & "'"
    Session("sljk_time")=now()
end if
if rs("mvalue")<30000 then
Response.Write "<script language=JavaScript>{alert('你月积分没有30000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("等级")>100 then
Response.Write "<script language=JavaScript>{alert('你等级大于[100]级不能使用申请哦！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("转生")>1 then
Response.Write "<script language=JavaScript>{alert('你是转生人不能申请金卡！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
sljk="爱情江湖站长看到[##]泡点的辛苦，已经连续在线<font color=red>"&mycd&"</font>分钟了，因此送他<img src='picwords/2.gif'>个金卡作为奖励,月积分减少<img src='picwords/2.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'>点！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>nothing
end function
%>