<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then 
	Response.Write "<script Language=Javascript>alert('提示：你不是正站长，你不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id=trim(request.form("id"))
a=trim(request.form("a"))
c=abs(int(request.form("c")))
d=abs(int(request.form("d")))
e=abs(int(request.form("e")))
f=trim(request.form("f"))
g=trim(request.form("g"))
if (c+d)>10000 then 
	Response.Write "<script language=javascript>{alert('提示：内力与武功的和不可以大于10000点!');parent.history.go(-1);}</script>" 
	Response.End 
end if 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sqlstr="SELECT * FROM n where id="&id
rs.Open sqlstr,conn,1,2
rs("a")=a
rs("c")=c
rs("d")=d
rs("e")=e
rs("f")=f
rs("g")=g
rs.Update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "managexx.asp?wgid="&id
%>