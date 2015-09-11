<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
regjm=Request.form("regjm")
regjm1=Request.form("regjm1")
if regjm<>regjm1 then
	Response.Write "<script Language=Javascript>alert('提示：你输入的认证码不正确，应该输入:"& regjm &"');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
name=request.form("name")
pass=trim(request.form("pass"))
if name="" or pass="" then
	Response.Write "<script Language=Javascript>alert('提示：是不是想开玩了？连大名和口令都不报？');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
pass=md5(pass)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="select id,密码 from 用户 where 姓名='" & name & "' and 密码='" & pass & "' "
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：有没有搞错？哪有这个人啊？');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if trim(pass)<>rs("密码") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：密码不对啊，会不会记错了？');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
id=rs("id")
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：请您记住自己的ID["&id&"]');window.close();</script>"
%>