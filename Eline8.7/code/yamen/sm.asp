<%@ LANGUAGE=VBScript codepage ="936" %>
<%sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="Select  * from �û� where ����='"&sjjh_name&"'"
set rs=conn.Execute(sql)
%>
<link href="../hc3w_img/hc3w_main.css" rel="stylesheet" type="text/css">
<body bgcolor="#FFFFFF">
<table border="0" cellSpacing="0" width="279" background="../hc3w_img/manager_bg.gif" height="388">
<TBODY>
  <tr>
    <td colSpan="2" width="100%" height="42"><table border="0" cellSpacing="1" width="100%">
<TBODY>
      <tr>
        <td width="20" align="center"><a href="sm.asp" target="gr"><img alt="===========&#13&#10���ɣ�<%=rs("����")%> &#13&#10��ݣ�<%=rs("���")%> &#13&#10������<%=rs("����")%>&#13&#10===========&#13&#10   ���ˢ�±�ҳ&#13&#10==========="  src='../ico/<%=rs("����ͷ��")%>-2.gif' border="0"></a></td>
        <td align="center"><%=rs("����")%> [<%=rs("���")%>] <b><font color="#dd2222"><%=rs("����")%></font></b> ��������<br></td>
      </tr>
</TBODY>
    </table>
    </td>
  </tr>
  <tr>
    <td colSpan="2" width="100%" height="81"><div align="center"><center><table border="0"
    cellPadding="0" cellSpacing="0" width="100%">
<TBODY>
      <tr>
        <td align="center">����</td>
        <td align="center">����</td>
        <td align="center">�书</td>
        <td align="center">�ڹ�</td>
        <td align="center">����</td>
      </tr>
      <tr>
        <td align="center"><img height="46" src="../hc3w_img/usergj.gif" width="47"></td>
        <td align="center"><img height="46" src="../hc3w_img/userfy.gif" width="47"></td>
        <td align="center"><img height="46" src="../hc3w_img/usersf.gif" width="47"></td>
        <td align="center"><img height="46" src="../hc3w_img/usertl_.gif" width="47"></td>
        <td align="center"><img height="46" src="../hc3w_img/useryz.gif" width="47"></td>
      </tr>
      <tr>
        <td align="center"><%=rs("����")%></td>
        <td align="center"><%=rs("����")%></td>
        <td align="center"><%=rs("�书")%></td>
        <td align="center"><%=rs("����")%></td>
        <td align="center"><%=rs("����")%></td>
      </tr>
</TBODY>
    </table>
    </center></div></td>
  </tr>
  <tr>
    <td colSpan="2" height="17"></td>
  </tr>
  <tr>
    <td height="17">&nbsp; ��&nbsp; �<%=rs("���")%> ��</td>
    <td height="17">��&nbsp; �£�<%=rs("��ż")%></td>
  </tr>
  <tr>
    <td height="17">&nbsp; ��&nbsp; �ѣ�<%=rs("��Ա��")%> Ԫ</td>
    <td height="17">��&nbsp; �£�<%=rs("����")%></td>
  </tr>
  <tr>
    <td height="17">&nbsp; ��&nbsp; ����<%=rs("����")%></td>
    <td height="17">Ӯ/�ģ�<%=rs("Ӯ����")%> / <%=rs("�Ĵ���")%></td>
  </tr>
  <tr>
    <td height="17">&nbsp; ��&nbsp; ����<%=rs("����")%></td>
    <td height="17">ս������<%=rs("�ȼ�")%> ��</td>
  </tr>
  <tr>
    <td height="6">&nbsp; �ĳ��ȼ���<%if rs("ӮǮ")<100  then response.write("�ĳ�����")
if rs("ӮǮ")>=2000 and rs("ӮǮ")<10000  then response.write("�ĳ�����")
if rs("ӮǮ")>=10000 and rs("ӮǮ")<50000  then response.write("������ͽ")
if rs("ӮǮ")>=50000 and rs("ӮǮ")<100000  then response.write("�ĳ�����")
if rs("ӮǮ")>=100000 and rs("ӮǮ")<300000  then response.write("ְҵ����")
if rs("ӮǮ")>=300000 and rs("ӮǮ")<800000  then response.write("����")
if rs("ӮǮ")>=800000 and rs("ӮǮ")<1500000  then response.write("����")
if rs("ӮǮ")>=1000000 then response.write("����")
%></td>
    <td height="6">������<%=rs("grade")%> ��</td>
  </tr>
</TBODY>
  <tr>
    <td height="6">&nbsp; �������ţ�<%if rs("�ȼ�")<5  then response.write("����է��")
if rs("�ȼ�")>=5 and rs("�ȼ�")<10  then response.write("��������")
if rs("�ȼ�")>=10 and rs("�ȼ�")<15  then response.write("С�гɾ�")
if rs("�ȼ�")>=15 and rs("�ȼ�")<20  then response.write("�����Ժ�")
if rs("�ȼ�")>=20 and rs("�ȼ�")<35  then response.write("�д�����")
if rs("�ȼ�")>=35 and rs("�ȼ�")<45  then response.write("һ������")
if rs("�ȼ�")>=45 and rs("�ȼ�")<55  then response.write("��������")
if rs("�ȼ�")>=55 and rs("�ȼ�")<65  then response.write("��������")
if rs("�ȼ�")>=65 and rs("�ȼ�")<75  then response.write("������ʤ")
if rs("�ȼ�")>=100 then response.write("����")
%></td>
    <td height="6">�����ˣ�<%=rs("������")%></td>
  </tr>
  <tr>
    <td height="5">&nbsp; E-MAIL��<%=rs("����")%></td>
    <td height="5">OICQ��<%=rs("oicq")%></td>
  </tr>
  <tr>
    <td height="5">&nbsp; ְ&nbsp; ҵ��<%=rs("ְҵ")%> ����:<%=rs("����")%></td>
    <td height="5">ʦ&nbsp; ����<%=rs("ʦ��")%></td>
  </tr>
  <tr>
    <td width="100%" colspan="2"><table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" align="center" height="30">����ǩ��</td>
  </tr>
  <tr>
    <td width="100%"><%=rs("ǩ��")%></td>
  </tr>
</table></td>
  </tr>
</table>

<p>
  <%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</p>