<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if Trim(Request.QueryString("w"))="save" then
 name=Trim(Request.Form("name"))
 url=Trim(Request.Form("url"))
 u_t=lcase(right(url,3))
 if name="" or url="" then
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
 Set conn=Server.CreateObject("ADODB.CONNECTION")
 conn.open Application("aqjh_usermdb")
 set rs=server.CreateObject ("ADODB.RecordSet")
 rs.open "song",conn,2,3
 rs.addnew
 rs("name")=name
 rs("url")=url
 rs.update
 rs.close
 response.redirect "song.asp"
end if 
%>
<html>
<head>
<title>����¸���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<table border=1 cellspacing=0 cellpadding=0 align=center width=680 bordercolorlight=#000000 bordercolordark=#FFFFFF>
<form action="song_add.asp?w=save" method=POST>
<tr> 
<td height=30 align=center colspan=2><b>����¸���</b></td></tr>
<tr> 
<td align=right width=11%>�������ƣ�</td>
<td width=*><input name="name" value="" size=30></td>
</tr>
<tr>
<td align=right>����URL��</td>
<td><input name="url" value="http://" size=95></td>
</tr>
<tr>
<td height=30 align=center colspan=2>
<input type="SUBMIT" name="Submit" value="���">
<input type="RESET" name="Reset" value="����">
</td>
</tr>
</form>
</table>
</body>
</html>