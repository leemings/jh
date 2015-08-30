<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'查看状态♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：查看状态需要管理员才可以操作！');}</script>"
	Response.End
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
'对系统禁止字符处理
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【查看资料】<font color=" & saycolor & ">管理员查看资料:"+getstat(towho)+"</font>"
towhoway=1
towho=sjjh_name
if says<>"" then
	call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
end if

'查看状态
function getstat(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & to1 &"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你所查看的人不存在！');}</script>"
	Response.End
end if
getstat=to1&":"& rs("性别") & ",伴侣:" & rs("配偶") & ",情人:" & rs("情人") & ",门派:" &rs("门派") & ",身份:" &rs("身份") &  ",武功:" &rs("武功") & ",内力:" & rs("内力") & ",体力:" & rs("体力") & ",攻击:" & rs("攻击") & ",防御:" & rs("防御")&",ip:" & rs("lastip") & ",状态:" & rs("状态") & ",信箱:" & rs("信箱") & ",魅力:" & rs("魅力") & ",赌次数:" & rs("赌次数") & ",赢次数:" & rs("赢次数") & ",赢钱:" & rs("赢钱") & ",道德:" & rs("道德") & ",管理等级:" & rs("grade") & ",总积分:" & rs("allvalue") & ",月积分:" & rs("mvalue") & ",登录次数:" & rs("times") & ",金:" & rs("金") & ",木:" & rs("木") & ",水:" & rs("水") & ",火:" & rs("火") & ",土:" & rs("土") & ",金币:" & rs("金币") & ",银币:" & rs("银币") & ",金卡:" & rs("会员金卡") & ",存款:" & rs("存款") & ",银两:" & rs("银两")
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
