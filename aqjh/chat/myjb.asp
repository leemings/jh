<%@ LANGUAGE=VBScript codepage ="936" %> 
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Weekday(date())=1 and Hour(time())>=20 and Hour(time())<21 then
	Response.Write "<script language=JavaScript>{alert('��ʾ������Ϊ����ʱ�䣬�������Բ���!');window.close();}</script>"
	Response.End 
end if
wpname=LCase(trim(Request.QueryString("wpname")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM b where a='"&wpname&"' and (b='ҩƷ' or b='�ʻ�')  order  by b,o",conn,2,2
if rs("n")="" or isnull(rs("n")) or rs("n")="��" then
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&wpname&"]���ڻ�û�о�Ӫ��!');window.close();}</script>"
	Response.End 
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%>
<html>
<script>
window.moveTo(100,30);
</script>
<head>
<title>��Ʒ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000" background="bg.gif">
<form method=POST action="myjbok.asp" name="form">
<table width="380" border="1" bordercolor="#000000" cellspacing="3" cellpadding="3" height="291" align="center">
  <tr> 
    <td bgcolor="#3399CC" height="2" colspan="4"> 
      <div align="center"><b><font color="#FFFF00">������Ʒ[<font color=blue><%=wpname%></font>](</font></b><font color="#FFFFFF" size="-1">��ʱֻ֧��ҩƷ���ʻ�</font><b><font color="#FFFF00">)</font></b></div>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#3399CC" height="23" width="21%">��Ӫ�ߣ�</td>
    <td bgcolor="#3399CC" height="23" width="25%"><b><font color="#FFFF00"><font color=blue><%=rs("n")%></font></font></b>
        <input type=hidden name=wpname value="<%=rs("a")%>" size="8" maxlength="10">
    </td>
    <td bgcolor="#3399CC" height="23" width="24%">Ͷ����:</td>
    <td bgcolor="#3399CC" height="23" width="30%"><b><font color="#FFFF00"><font color=blue><%=rs("q")%></font></font></b>��</td>
  </tr>
  <tr> 
    <td bgcolor="#3399CC" width="21%" height="26">��ۣ�</td>
    <td bgcolor="#3399CC" width="25%" height="26"> 
      <input type=text name=money value="<%=rs("h")%>" size="8" maxlength="10" readonly>
    </td>
    <td bgcolor="#3399CC" width="24%" height="26">���룺</td>
    <td bgcolor="#3399CC" width="30%" height="26"><b><font color="#FFFF00"><font color=blue><%=rs("o")%></font></font></b>��</td>
  </tr>
  <tr> 
    <td bgcolor="#3399CC" colspan="4" height="19"> 
      <div align="center">��Ӫ�߹���</div>
    </td>
  </tr>
  <tr align="left" valign="top"> 
    <td bgcolor="#3399CC" colspan="4" height="118">
      <div align="center"><b><font color="#FFFF00"><font color=blue> 
        <textarea name="ggsm" <%if rs("n")<>aqjh_name and aqjh_grade<10 then%> readonly <%end if%> cols="45" rows="5"><%=rs("p")%></textarea>
        <br>
        </font></font></b><font color="blue" size="-1">��Ϊ��Ӫ�ߣ�����������޸���ۼ����棡</font></div>
    </td>
  </tr>
  <%if rs("n")=aqjh_name then%>
  <tr> 
    <td bgcolor="#3399CC" colspan="4"> 
      <div align="center">
        <input type=submit value=ȷ���޸� name="submit" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
      </div>
    </td>
  </tr>
 <%end if%> 
 <%
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing%>
</table>
</form>
</body>
</html>
