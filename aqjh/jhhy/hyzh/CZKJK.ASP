<%@ LANGUAGE=VBScript codepage ="936" %>
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
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<title>充值会员金卡</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="bg.gif">
<%my=aqjh_name
input=request.form("input1")
if not isnumeric(input) then 
	Response.Write "<script language=JavaScript>{alert('提示：["&input&"]输入错误，请使用数字！');location.href = 'czk2.asp';}</script>"
	Response.End 
end if
input=int(abs(input))
if input/1<>int(input/1) then
 	Response.Write "<script language=JavaScript>{alert('请输入1的整数倍的数值！');location.href = 'czk2.asp';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 充值卡 from 用户 where 姓名='" & aqjh_name & "'",conn,2,2
if rs("充值卡")<input then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('提示：你没有这么多的充值卡无法充值！');location.href = 'czk2.asp';}</script>"
	response.end
end if
conn.execute "update 用户 set 充值卡=充值卡-"&input&",会员金卡=会员金卡+" & int(input*5) & " where 姓名='"&aqjh_name&"'"
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('"&aqjh_name&"把"&input&"张充值卡充值成会员金卡:"& int(input*5)&"张！');location.href = 'czk2.asp';</script>"
%>
