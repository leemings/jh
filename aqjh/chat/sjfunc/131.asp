<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'申领金币
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
says="<font color=red>【申领金币】<font color=" & saycolor & ">"+sljb()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function sljb()
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if Session("sljb_time")="" then Session("sljb_time")=rs("lasttime")
mycd=DateDiff("n",Session("sljb_time"),now())
jbmoney=2
hyjb=2
if mycd<60 then
    if DateDiff("n",rs("lasttime"),now())>=60 then
       sljb="你在一个小时前已经领过了"
    else
       sljb="你现在泡江湖时间还不到一小时(还差:"& (60-mycd) &"分到1小时),如果您连续泡点1小时以上(退出,掉线,打死时间清0)可以领取"& jbmoney &"个金币,付费会员是普通玩家"& hyjb &"倍!"
    end if
    Response.Write "<script language=JavaScript>{alert('"&sljb&"');}</script>"
    Response.End
end if
hy=rs("会员等级")
mycd=int(mycd/60)
if rs("会员等级")=4 then
	myjbsl=jbmoney*hyjb*mycd
        conn.execute "update 用户 set 金币=金币+"& myjbsl &" where 姓名='" & aqjh_name & "'"
        Session("sljb_time")=now()
	sljb="##为<font color=blue>"& hy &"</font>级会员，乃付费会员，在江湖努力泡点<font color=red><b>"& mycd &"</b></font>小时,得到金币:<font color=red>"& myjbsl &"</font>个,##会继续努力,支持快乐江湖的发展!"
else
	myjbsl1=jbmoney*mycd
        conn.execute "update 用户 set 金币=金币+"& myjbsl1 &" where 姓名='" & aqjh_name & "'"
        Session("sljb_time")=now()
	sljb="##为<font color=blue>"& hy &"</font>级会员，乃免费会员，在江湖努力泡点<font color=red><b>"& mycd &"</b></font>小时,得到金币:<font color=red>"& myjbsl1 &"</font>个,##会继续努力,支持快乐江湖的发展!"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>