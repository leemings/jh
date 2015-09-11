<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 姓名,会员等级,性别,密码 from [用户] WHERE 姓名='" & Session("aqjh_name") & "'",conn,1,1
hydj=rs("会员等级")
pass=rs("密码")
if rs("性别")="男" then
	mysex="M"
	myshow="||||||14|13|12||11||10|9||||8|||||||"
else
	mysex="F"
	myshow="||||||7|6|5||4||3|2||||1|||||||"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from use where a='"& Session("aqjh_name") &"'", conn, 1,3
if rs.eof and rs.bof then
rs.addnew
end if
rs("a")=Session("aqjh_name")
rs("b")=pass
rs("c")=mysex
if hydj<>0 then
	rs("d")=true
else
	rs("d")=false
end if
rs("e")=200
rs("f")=myshow
rs("g")=""
rs("h")=now()
rs("i")=0
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "main.asp"
%>
