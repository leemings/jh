<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
money=clng(request.form("money"))
if not isnumeric(input) then 
	Response.Write "<script language=JavaScript>{alert('提示：[" & input & "]输入错误，请使用数字！');location.href = 'hy.asp';}</script>"
	Response.End 
end if
input=int(abs(money))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 金币 from [用户] where 姓名='" & aqjh_name & "'",conn,1,1
hyjb=rs("金币")
if rs("金币")<input then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你没有这么多的金币无法转换！');location.href = 'hy.asp';}</script>"
	response.end
end if
conn.execute "update [用户] set 金币=金币-" & input & " where 姓名='" & aqjh_name & "'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
conn.execute "update use set e=e+" & input*10 & " where a='" & aqjh_name & "'"
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('" & aqjh_name & "您转换金币:" & input*10 & "个！');location.href = 'hy.asp';</script>"
%>
