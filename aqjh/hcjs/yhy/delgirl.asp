<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ID,银两 FROM 用户 WHERE 姓名='" & aqjh_name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('江湖中找不到你的资料！');window.close();</script>"
	response.end
end if
aqjh_id=rs("id")
yinliang=rs("银两")
rs.close
%><!--#include file="jiu.asp"--><%
sql="select * from 妓女 where 妓女ID="& aqjh_id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('你弄错了吧，本院没有你这样的姑娘呀！');window.close();</script>"
	response.end
end if
if yinliang<5000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('你的腰包里没有500万，想来赎身，门都没！');window.close();</script>"
	response.end
end if
conn.execute "update 用户 set 银两=银两-5000000 where 姓名='"& aqjh_name &"'"
connt.execute("delete * from 妓女 where 妓女ID="&aqjh_id)
rs.close
set rs=nothing
conn.close
set conn=nothing
connt.close
set connt=nothing
Response.Write "<script Language=Javascript>alert('这是你的卖身契，你现在自由了，欢迎你在没钱的时候在来我们这工作！');window.close();</script>"
response.end
%>