<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"


sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"


if sjjh_name="" then
	response.write("对不起，你还没有<a href=../../index.htm>登录</a>")
else
	sex=request.form("sex")
if sex="" then
	response.write("对不起，你还没有<a href=../../index.htm>登录</a>")
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 where 姓名='"&sjjh_name&"'"&" and 性别= '" &sex&"'",conn
if rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=436"
else
if rs("银两")<380 then
Response.Redirect "../../error.asp?id=437"
else
if day(rs("领钱"))<day(now()) or month(rs("领钱"))<month(now()) or year(rs("领钱"))<year(now()) or isnull(rs("领钱")) then
energy=rs("体力")
tempdate=date
energytemp=energy+1000
sql="update 用户 set 银两=银两-300,领钱='"&tempdate&"',体力='"&energytemp&"' where 姓名='"&sjjh_name&"'"
set rs=conn.execute(sql)
conn.close
set conn=nothing
set rs=nothing

Response.Redirect "../../ok.asp?id=601"
else
rs.close
Response.Redirect "../../error.asp?id=438"
end if
end if
end if
end if
end if
%>
