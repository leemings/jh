<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
if sjjh_name<>Application("sjjh_user") then 
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻����վ�����㲻�ܲ�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
wgid=trim(request.querystring("wgid"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
	rs.open "Select * From y Where id="&wgid,conn
	if rs.eof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Redirect "../error.asp?id=447"
	else
		conn.execute "Delete * From y Where id="&wgid
		conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','ɾ���书��["&wgid&"]')"
		Response.Redirect "../ok.asp?id=704"
	end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script language=vbscript>
MsgBox "<%=response.write(str)%>"
location.href = "adminwg.asp?mp=<%=rs("b")%>"
</script>
