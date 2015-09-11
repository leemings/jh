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
userid=request.form("userid")
pass=trim(request.form("pass"))
if name="" or pass="" then
	Response.Write "<script Language=Javascript>alert('提示：请输入用户名和保护密码！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
pass=md5(pass)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="select 第二密码 from 用户 where id="&userid&" and 姓名='" & name & "' and 第二密码='" & pass & "' "
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：密码保护出错！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if trim(pass)<>rs("第二密码") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：保护密码不对啊，会不会记错了？');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
sql="update 用户 set 密码='e10adc3949ba59abbe56e057f20f883e' where 姓名='"&name&"'"
conn.Execute(sql)
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：请记住您的密码[123456]并尽快改掉！');window.close();</script>"
%>