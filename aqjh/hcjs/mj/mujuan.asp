<%@ codepage ="936" %>
<%
username=session("aqjh_name")
if username="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
const MaxPerPage=10
sql="select * from �û�  where ����='"&username&"'"
set rs=conn.execute(sql)
%>
<html>
<head>
<title>����ļ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../css.css">
</head>
<body bgcolor="#000000" text="#FFFFFF" background="../../bg.gif">
<p align="center"><img border="0" src="../../chat/img/menoy.gif"><b><font size="4" color="#008080">����ļ��<img border="0" src="../../chat/img/1414.jpg"></font></b></p>
<form method="post" action="mj.asp">
  <table width="500" border="1" cellspacing="0" cellpadding="5" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr> 
      <td bgcolor="#008080"><font color="#FFFFFF">ļ���ˣ�</font></td>
      <td bgcolor="#008080"><font color="#FFFFFF"><%=session("aqjh_name")%>��</font></td>
      <td bgcolor="#008080"><font color="#FFFFFF">����:<%=rs("����")%></font></td>
    </tr>
    <tr> 
      <td bgcolor="#008080"><font color="#FFFFFF">����������</font></td>
      <td bgcolor="#008080"><font color="#FFFFFF"><%=rs("����")%>&nbsp;Ԫ</font></td>
      <td bgcolor="#008080"><font color="#FFFFFF">����:<%=rs("����")%></font></td>
    </tr>
    <tr> 
      <td bgcolor="#008080"><font color="#FFFFFF">ļ����:</font></td>
      <td bgcolor="#008080"> 
        <font color="#FFFFFF"> 
        <input type="text" name="mj" size="10" maxlength="10">
        ��С���10000</font></td> 
      <td bgcolor="#008080"><font color="#FFFFFF">����:<%=rs("����")%></font></td>
    </tr>
    <tr> 
      <td colspan="2" bgcolor="#008080" > 
        <div align="center"> 
          <font color="#FFFFFF"> 
          <input type="submit" name="Submit" value="ļ ��">  ��   
          <input type="reset" name="Submit2" value="�� ��"></font>
        </div>
      </td>
      <td bgcolor="#008080"><font color="#FFFFFF">����:<%=rs("����")%></font></td>
    </tr>
  </table>
</form>
<div align="center"><font color="#000000"><br>
  </font> 
  <font size="2" color="#339966">����ļ��ɶ����д��ĺô�Ӵ��Ǯ����û�ط������Ծ���������<br>
  ��100000�������������������10,�������10,���»�����20Ӵ<br>
  û�������ɲ��ܽ��ѽ��������Ը���Ҹ������������?? 
  </font></body></html>