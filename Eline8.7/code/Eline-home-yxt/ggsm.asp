<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
if sjjh_name<>Application("sjjh_user") then 
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻����վ�����㲻�ܲ�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
cz=trim(request.querystring("cz"))
if cz<>"����" then cz=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM sm where a='���"&cz&"'",conn,2,2
%>
<html>
<head>
<title>�����Ҷ�������޸ġ�wWw.happyjh.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center">�����Ҷ�������޸�(&lt;br&gt;Ϊ�س� ֧��html����)
<p><br>
  <font size="+2" face="����_GB2312">��������޸Ĵ���,�����һЩ�벻���ĺ����������벻�������ҵȣ�������޸Ĵ������<a href="ggsm.asp?cz=%B1%B8%B7%DD"><b><font color="#FF0000">����</font></b></a>���б�Ҫ���������ٵ��޸ľͿ��Իָ�����ʼ����״̬ 
  ��ע���޸Ĳ�Ҫ����!</font>
<p align="center">(<font color="#0000FF">�����Ա˵��</font>) <br>

<form method=POST action="ggsmok.asp?cz=���" name="form">
  <div align="center"> ����ȣ� 
    <input type="text" name="hydz" size="10" value="<%=rs("b")%>" maxlength="2">
    (����Ϊ0��رչ��)<br>
    <textarea name="hysm"  cols="90" rows="14"><%=rs("c")%></textarea>
    <br>
    <br>
    <br>
    <input type=submit value=ȷ����������޸� name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
  </div>
</form>
<%rs.close
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
