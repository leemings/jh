<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from use where a='"& aqjh_name &"'", conn, 1,3
if rs.eof and rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "reg.asp"
	Response.End
end if
if rs("c")="M" then
	rs("f")="||||||14|13|12||11||10|9||||8|||||||"
else
	rs("f")="||||||7|6|5||4||3|2||||1|||||||"
end if
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：数据库中形像初始化成功！');window.close();}</script>"
Response.End
%>