<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM sm where a='��Ա'",conn,2,2
%>
<html>
<head>
<title>�������ɼ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="img/snowbg.gif">
<div align="center">
  <p><b><font color="#0000CC" size="+2">�ǵġ�������ά��������������<br>
    </font></b></p>
  <p align="left">1.���ɼ��˵�������<br>
    a.����һ����ģ�����ɡ�<br>
    <br>
    b.�����ṩ������������.<br>
    <br>
    c.�ڱ������߱�һ��Ӱ����������.<br>
    <br>
    d.����������ά�������ṩ����ע�����ϣ�<br>
    <br>
    e.����IP��ַ���̶��ߣ��������޸�IP���á�<br>
    <br>
    2.�����շѱ�׼���Ż����ߣ�<br>
    <br>
    a.���������շѣ�200Ԫ/��(�ۡ��¡�̨��������������Ԫ����200$Ԫ/��!)<br>
    (<b><font color="#FF0000">�����ڼ��ؼ�:150/�£��ۡ��¡�̨���䣡</font></b>) <br>
    <br>
    b.�ڼ�����������ÿ����Զ���ȡ����10����!<br>
    <br>
    c.�ڼ������������������״̬���ᶪʧ��<br>
    <br>
    d.�ڼ�������������<b><font color="#0000FF">�ݵ���ƽ������������2���� </font></b><br>
    <br>
    e.������ά�����������ڶ����������֧��һ�ܣ�<br>
    <br>
    f.�ڼ����������������7�ۣ�(��Ա����)<br>
    <br>
    <br>
    <br>
    3.���˰취��<br>
  </p>
  <table width="75%" border="1" align="left">
    <tr>
      <td bgcolor="#0066CC"><%=rs("d")%> 
        <%rs.close
set rs=nothing
conn.close
set conn=nothing%>
      </td>
    </tr>
  </table>
  <p align="left"> <br>
  </p>
</div>
</body>
</html>
