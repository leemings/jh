<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then 
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻����վ�����㲻�ܲ�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id=clng(Request("id"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
'����
if Trim(Request.QueryString("w"))="save" then
 name=Trim(Request.Form("name"))
 url=Trim(Request.Form("url"))
 u_t=lcase(right(url,3))
 if url="" or name="" then
%>
<script language=javascript>
     history.back()
     alert("����:��д���ݲ���Ϊ�գ�")
</script>
<%
 respoonse.end
 end if
 if u_t<>"mp3" and u_t<>"wma" and u_t<>"mid" then
	response.write "�������Ͳ���ȷ����ʹ��mp3|wma|mid��ʽ"
	response.end
 end if
 Set rs=Server.CreateObject("ADODB.RecordSet")
 rs.open "select * from song where id="&request("id"),conn,2,3
 rs("name")=name
 rs("url")=url
 rs.update
 rs.close
 response.redirect "song.asp"
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.open "select * from song where id="&id,conn,1,3
if rs.eof then
Response.Write ("��ʾ:�����д�...")
response.end
end if
%>
<html>
<head>
<title>��̨��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<table width=680 border=1 cellspacing=0 cellpadding=3 bordercolorlight=#000000 bordercolordark=#FFFFFF align=center>
<tr align=center><td height=30 align=center colspan=2><b>�޸ĸ���</b></td></tr>
<form action="song_edit.asp?w=save&id=<%=rs("id")%>" method=POST>
<tr>
<td align=right width=11%>�����ɣģ�</td>
<td><input name="id" type="text" size=10 maxlength=5 value="<%=RS("id")%>" readonly>
</tr>
<tr>
<td align=right width=11%>�������ƣ�</td>
<td width=*><input name="name" size=30 value="<%=RS("name")%>"></td></tr>
<tr> 
<td align=right>����URL��</td>
<td><input name="url" value="<%=RS("url")%>" size=95></td>
</tr>
<tr><td align=center colspan=2>
<input type="SUBMIT" name="Submit" value="����">
<input type="RESET" name="Reset" value="����">
</td>
</tr>
</form>
</table>
</body>
</html>
<%
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>