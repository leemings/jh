<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<%
cz=trim(request.querystring("cz"))
if cz<>"����" then cz=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<html>
<head>
<title>�����Ҷ�������޸�</title><LINK href="css/css.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center">�����Ҷ�������޸�(&lt;br&gt;Ϊ�س� ֧��html����)
<p><br>
��������޸Ĵ���,�����һЩ�벻���ĺ����������벻�������ҵȣ�������޸Ĵ������<a href="ggsm.asp?cz=%B1%B8%B7%DD"><b><font color="#FF0000">����</font></b></a>���б�Ҫ���������ٵ��޸ľͿ��Իָ�����ʼ����״̬ 
  ��ע���޸Ĳ�Ҫ����!</font>
<%
rs.open "SELECT * FROM sm where a='����"&cz&"'",conn,2,2%>
<form method=POST action="ggsmok.asp?cz=����" name="form">
  <div align="center"><font color="#0000FF"><b>���뽭���Ĺ������</b></font><br>
    (�ǻ�Ա��ÿһ��֮����;�ֺŰ�Ƿָ�)<br>
    <textarea name="hysm"  cols="90" rows="8"><%=rs("c")%></textarea>
    <br>
    (��Ա����˵����ֻ��һ�仰\n��ʾ����!) <br>
    <textarea name="hydz"  cols="90" rows="4"><%=rs("d")%></textarea>
    <br>
    <br>
    <br>
    <input type=submit value=ȷ���������޸� name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
  </div>
</form>
<%rs.close
rs.open "SELECT * FROM sm where a='pkֵ��'",conn,2,2%>
<form method=POST action="ggsmok.asp?cz=pkֵ��" name="form">
  <div align="center"><font color="#0000FF"><b>�ٸ�ֵ�ձ�</b></font><br>
    <textarea name="hysm"  cols="90" rows="8"><%=rs("c")%></textarea>
    <br>
    <input type=submit value=ȷ���޸�ֵ�ձ� name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
  </div>
</form>
<%rs.close
rs.open "SELECT * FROM sm where a='����"&cz&"'",conn,2,2%>
<form method=POST action="ggsmok.asp?cz=����" name="form">
  <div align="center"><font color="#0000FF"><b>��������������������</b></font><br>
    ÿһ��֮����;�ֺŰ�Ƿָ�<br>
    <textarea name="hysm"  cols="90" rows="12"><%=rs("c")%></textarea>
    <br>
    ������ʿ��ʽ���£���1|��1;��2|��2;<br>
    <textarea name="hydz"  cols="90" rows="12"><%=rs("d")%></textarea>
    <br>
    <br>
    <br>
    <input type=submit value=ȷ����������޸� name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
  </div>
</form>
<%rs.close
set rs=nothing
conn.close
set conn=nothing%> </body>
</html>