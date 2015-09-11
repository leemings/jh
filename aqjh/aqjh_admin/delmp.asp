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
owner=trim(request.querystring("rs"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
	rs.open "Select * From p Where a='" & owner & "'",conn
	if rs.eof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Redirect "../error.asp?id=447"
	else
		conn.execute "Delete * From p Where a= '" & owner & "' "
		conn.execute("Update 用户 set 门派='游侠',身份='弟子',grade=1 where 门派='"&owner&"'")
		conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','删除门派：["&owner&"]')"
		Response.Redirect "../ok.asp?id=704"
	end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script language=vbscript>
MsgBox "<%=response.write(str)%>"
location.href = "adminmpp.asp"
</script>