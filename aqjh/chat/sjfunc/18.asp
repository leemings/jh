<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'查看状态
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_grade<grade_see then
	Response.Write "<script language=JavaScript>{alert('提示：查看状态需要管理等级["&grade_see&"]才可以操作！');}</script>"
	Response.End
end if
call dianzan(towho)
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
if len(says)>1500 then says=Left(says,1500)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=1
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【状态】<font color=" & saycolor & ">##得到江湖情报："+getstat(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'查看状态
function getstat(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & to1 &"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你所查看的人不存在！！！');}</script>"
	Response.End
end if
if to1=Application("aqjh_user") then
    	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你想查总站长,找死啊！！！');}</script>"
	Response.End
end if
getstat=to1 &"，" & rs("性别") & "，伴侣：" & rs("配偶") & "，情人：" & rs("情人") & "，师傅：" & rs("师傅") & "，门派：" &rs("门派") & "，身份：" &rs("身份") &  "，武功：" &rs("武功") & "，内力：" & rs("内力") & "，体力：" & rs("体力") & "，攻击力：" & rs("攻击") & "，防御力：" & rs("防御")&"，登录ip：" & rs("lastip") & "，信箱：" & rs("信箱") & "，魅力：" & rs("魅力") & "，赌次数：" & rs("赌次数") & "，赢次数：" & rs("赢次数") & "，赢钱：" & rs("赢钱") & "，道德：" & rs("道德") & "，管理等级：" & rs("grade") & "，累积分：" & rs("allvalue") & "，月积分：" & rs("mvalue") & "，登录：" & rs("登录") & "，银两：" & rs("银两")&"，金币："&rs("金币")&"，存款："&rs("存款")&"，介绍人："&rs("介绍人")
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>