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
<title>����ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: ����}
input {  font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px}
.c {  font-family: ����; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body >
<p>&nbsp;</p>
<p align="center"><b><font size="3">����</font></b><b><font size="3">��������ϵͳ</font></b></p>
<form action="jiangli2.asp" method=POST >
  <div align="center"> ����  
    <select name="dj">
      <option value="��Ա�ȼ�<1" selected>�ǻ�Ա</option>
      <option value="��Ա�ȼ�>0">��Ա</option>
      <option value="grade=5">����</option>
      <option value="grade>6">�ٸ�</option>
    </select>
    <input type="text" name="sl" maxlength="2" size="5" value="10"> 
    �� ,  
    <input type="text" name="number" value="100" maxlength="5" size="10">
    <select name="xm"> 
      <option value="allvalue" selected>�ݵ�</option>
      <option value="���">���</option>
      <option value="�ݶ�����">����</option>
      <option value="��Ա��">��</option>
      <option value="����">����</option>
      <option value="����">����</option>
      <option value="����">����</option>
      <option value="����">����</option>
      <option value="����">����</option>
      <option value="�书">�书</option>
      <option value="����">����</option>
      <option value="����">����</option>
      <option value="���">���</option>
      <option value="����">����</option>
      <option value="�Ṧ">�Ṧ</option>
    </select>
    <input type="submit" name="Submit" value="����"></div> 
</form>
</body>
</html>
