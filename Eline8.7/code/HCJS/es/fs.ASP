<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
sl=abs(int(Request.form("wpsl")))
fs=int(Request.form("wpfs"))
id=trim(request.querystring("id"))
if not isnumeric(id) then Response.Redirect "../../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select i,a from w where id=" & id & " and i>0",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你没有这种物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if int(rs("i"))<sl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：操作错误，["&rs("a")&"]只有"&rs("i")&"个物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
select case fs
case 1
	Response.Redirect "zs.asp?id="& id &"&sl="& sl
case 2
	Response.Redirect "jy.asp?id="& id &"&sl="& sl
case 3
	Response.Redirect "es.asp?id="& id &"&sl="& sl
case 4
	Response.Redirect "bxg.asp?id="& id &"&sl="& sl
end select
%>