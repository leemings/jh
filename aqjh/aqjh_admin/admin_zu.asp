<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>

<title>��ӹ���</title>
<link rel="stylesheet" href="css/css.css" type=text/css>
</head>

<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<div align="center"> 
  <p><br>
    <font color="#009900"><b><font color="#FF0000" size="+6">[��ӹ���]</font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="700"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF" align="center"> 
    <td height="27" width="54" align="center"><font color="#FF6600">��ID</font>
    </td>
    <td height="27" width="55" align="center"><font color="#FF6600">�� ��</font>
    </td>
    <td height="27" width="40" align="center"><font color="#FF6600">�鳤ID</font>
    </td>
    <td height="27" width="120" align="center"><font color="#FF6600">�� Ա</font>
    </td>
    <td height="27" width="120" align="center"><font color="#FF6600">���ʱ��</font>
    </td>
    <td height="27" width="40" align="center"><font color="#FF6600">��Ա��</font>
    </td>
    <td height="27" width="66" align="center"><font color="#FF6600">��  ��</font>
    </td>
  </tr>
  <%Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
rs.open "Select * From ���",conn,0,1
if not rs.eof then
%>
  <%do while not rs.eof%>
  <form method=POST action='admin_zuok.asp?id=<%=rs("id")%>'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="54"> 
        <div align="center"><font color="#FF6600">ID:</font> <font color="#FF6600"><%=rs("id")%></font></div>
        <div align="center"></div>
      </td>
      <td height="2" width="55" align="center"> 
        <div align="center"> 
          <input type="text" name="zuname" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" value="<%=rs("����")%>" maxlength="255" readonly>
      </td>
      <td height="2" width="40" align="center">
          <input type="text" name="zuzhang" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="4"
value="<%=rs("�鳤")%>" maxlength="255" readonly>
      </td>
      <td height="2" width="120" align="center"> 
        <div align="center"> 
          <input type="text" name="zuyuan" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="30"
value="<%=rs("��Ա")%>" maxlength="255">
      </td>
      <td height="2" width="120" align="center"> 
        <div align="center"> 
          <input type="text" name="zudata" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="30"
value="<%=rs("ʱ��")%>" maxlength="255" readonly>
      </td>
      <td height="2" width="40" align="center"> 
        <div align="center"> 
          <input type="text" name="zucount" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="2" value="<%=rs("��Ա��")%>">
      </td>
      <td height="2" width="66"> 
          <input type="submit" value="�޸�"
name="submit">&nbsp;<input type="submit" value="ɾ��"
name="submit">
      </td>
    </tr>
  </form>
<%rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%>
</table>
</body>
</html>
