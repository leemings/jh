<%@ LANGUAGE=VBScript codepage ="936" %>

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
peiou=clng(trim(Request("peiou")))
guan=clng(trim(Request("guan")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 配偶 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,1,1
mywife=rs("配偶")
rs.close
if peiou<>0 then
	rs.open "SELECT 姓名 from [用户] WHERE 姓名='" & mywife & "'",conn,1,1
	if rs.Eof or rs.Bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：找不到你的配偶资料,没有办法让他进来!');javascript:history.go(-1);</script>"
		response.end
	end if
	mywifezh=rs("姓名")
	rs.close
else
	mywifezh="无"
end if
rs.open "SELECT h_拥有者2,h_参观 FROM house WHERE h_拥有者='" & aqjh_name &"'",conn,1,3
if rs.Eof or rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('提示：你还没有房屋或者你无权设置！');parent.email.style.visibility='hidden';</script>"
	Response.End
end if
rs("h_拥有者2")=mywifezh
rs("h_参观")=guan
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=""Javascript"">alert('提示：设置成功!');location.href = 'set.asp';</script>"
%>