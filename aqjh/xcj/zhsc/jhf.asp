<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
%>
<title>金卡转换法力</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../../bg.gif">
<%my=aqjh_name
input=request.form("input")
if not isnumeric(input) then 
	Response.Write "<script language=JavaScript>{alert('提示：["&input&"]输入错误，请使用数字！');location.href = 'flzh.asp';}</script>"
	Response.End 
end if
input=int(abs(input))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 会员金卡,会员等级,等级,法力,法力加 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
hy=rs("会员等级")
if rs("会员金卡")<input then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你没有这么多的金卡无法转换！');location.href = 'flzh.asp';}</script>"
	response.end
end if
if rs("等级")<18 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你等级不够18级不能转换！');location.href = 'flzh.asp';}</script>"
	response.end
end if
if input<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你想死啊，没钱还想换！');location.href = 'flzh.asp';}</script>"
	response.end
end if
if (rs("法力")+rs("法力加"))>20000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你总法力大于2000万，已经够多了！');location.href = 'flzh.asp';}</script>"
	response.end
end if
conn.execute "update 用户 set 会员金卡=会员金卡-" & input & ",法力=法力+"& int(input*5000)&" where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>【转换市场】["&aqjh_name&"]</font><font color=blue>把"&input&"元金卡转换成了"&int(input*5000)&"点法力,于是法力大增!</font>"
call showchat(says)
Response.Write "<script Language=Javascript>alert('"&aqjh_name&"您转换法力:"&input*5000&"点！');location.href = 'flzh.asp';</script>"
%>