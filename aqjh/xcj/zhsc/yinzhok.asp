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
<title>现金转换金币</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../../bg.gif">
<%my=aqjh_name
input=request.form("input1")
if not isnumeric(input) then 
	Response.Write "<script language=JavaScript>{alert('提示：["&input&"]输入错误，请使用数字！');location.href = 'yinzh.ASP';}</script>"
	Response.End 
end if
input=int(abs(input))
if input/1<>int(input/1) then
 	Response.Write "<script language=JavaScript>{alert('请输入1的整数倍的数值！');location.href = 'yule2.asp';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 金币,银两,存款,等级 from 用户 where 姓名='" & aqjh_name & "'",conn,2,2
if rs("金币")<input then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('提示：你没有这么多的金币无法转换！');location.href = 'yule2.asp';}</script>"
	response.end
end if
if rs("等级")<25 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你等级不够25级不能转换！');location.href = 'yinzh.asp';}</script>"
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
if (rs("银两")+rs("存款"))>2000000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你钱已经大于20亿了，已经够多了！');location.href = 'flzh.asp';}</script>"
	response.end
end if
conn.execute "update 用户 set 金币=金币-"&input&",银两=银两+" & int(input*1000000) & " where 姓名='"&aqjh_name&"'"
conn.close
set conn=nothing
says="<font color=red>【转换市场】["&aqjh_name&"]</font><font color=blue>把"&input&"个金币兑现成了"&int(input*1000000)&"银两，成大富翁了!</font>"
call showchat(says)
Response.Write "<script Language=Javascript>alert('"&aqjh_name&"把"&input&"个金币转换成现银两:"& int(input*1000000)&"两！');location.href = 'yinzh.asp';</script>"
%>
