<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
a=trim(request.form("a"))
e=trim(request.form("e"))
g=trim(request.form("g"))
d=trim(request.form("d"))
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"jhmp")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM p where b='" & aqjh_name & "' and a='"&a&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你好呀！黑客先生，这回不灵了吧？！');location.href = 'javascript:history.back()';</script>"
	response.end
end if
rs.close
sqlstr="SELECT * FROM p where a='"&a&"'"
rs.Open sqlstr,conn,1,2
rs("e")=e
rs("g")=g
rs("d")=d
rs.Update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "setwg.asp"
%>