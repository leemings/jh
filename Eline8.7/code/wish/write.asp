<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="config.asp"-->
<!--#include file="conn.asp"-->

<%
if Session("sjjh_name")=""  then response.end
nickname=Session("sjjh_name")
grade=Session("sjjh_grade")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="select ����,����,����,�Ա� from �û� where ����='"&Session("sjjh_name")&"'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%>
<script language="vbscript">
  MsgBox "���󽭺�û��������˰�"
  location.href = "../jh.asp"
</script>
<%
else
email=rs("����")
menpai=rs("����")
sex=rs("�Ա�")%>

<html>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"> 
<head>
<LINK 
href="../css.css" rel=stylesheet>
<title><%=title%>��wWw.51eline.com��</title>
</head>

<body bgcolor="<%=bgcolor%>" text="<%=textcolor%>" link="<%=linkcolor%>">

<form method="post" action="save.asp" align="center">
  <center>
    <table
    border="1" cellspacing="2" width="415" bordercolor="<%=titlelightcolor%>"
    height="310" cellpadding="2">
      <tr>
            
        <td align="center" colspan="2" bgcolor="<%=titledarkcolor%>"><font color="#FFFFFF">��<%=nickname%>�� 
          ����Ը��</font></td>
        </tr>
        <tr>
            <td colspan="2" height="23"><p align="center">&nbsp;<font
            color="<%=titlelightcolor%>">�� ��*�� ������д </font></p>
            </td>
        </tr>
        <tr>
            
        <td align="center" width="80" height="23">
<p
            align="left">���������� </p>
            </td>
            
        <td height="23" width="315">
<p align="left">
            <input type="text"
            size="15" maxlength="20" name="name" value="<%=rs("����")%>"readonly>
          </p>
            </td>
        </tr>

        <tr>
            
        <td align="center" width="80" height="24">
<p
            align="left">�����Ա� </p>
            </td>
            
        <td height="24" width="315"> 
          <p align="left">
            <input type="text"
            size="8" maxlength="20" name="sex" value="<%=sex%>"readonly>
          </p>
            </td>
        </tr>
        <tr>
            
        <td align="center" width="80" height="23">
<p
            align="left">�������䣺 </p>
            </td>
            
        <td height="23" width="315">
<p align="left">
            <input type="text"
            size="15" maxlength="50" name="email" value="<%=email%>"readonly>
          </p>
            </td>
        </tr>
        <tr>
            
        <td align="center" width="80" height="23"> 
          <p
            align="left">�������ɣ� </p>
            </td>
            
        <td height="23" width="315">
<p align="left">
            <input type="text"
            size="15" maxlength="50" name="homepage"
            value="<%=menpai%>"readonly>
             </p>
            </td>
        </tr>
        <tr>
            
        <td width="80" height="23">
<p align="left">Ը�����
            </p>
            </td>
            
        <td height="23" width="315">
<p align="left">
            <select name="wishtype"
            size="1">
              <option value="love">��������</option>
                <option value="study">ѧ����ҵ</option>
                <option value="health">��������</option>
                <option value="family">�ҡ���ͥ</option>
                <option value="work">�¡���ҵ</option>
                <option value="future">��������</option>
                <option value="wealth">�ơ�����</option>
                <option value="life">��������</option>
            </select> ��*�� </p>
            </td>
        </tr>
        <tr>
            
        <td width="80" class="pt9" height="80">��ס�ط���</td>
            
        <td class="pt9" width="315" height="80"> 
          <table border="0" width="300">
            <tr> 
              <td width="53"> 
                <input type="radio" name="address" value="����" checked>
                ����</td>
              <td width="70"> 
                <input type="radio" name="address" value="����">
                ���� </td>
              <td width="51"> 
                <input type="radio" name="address" value="����">
                ���� </td>
              <td width="53"> 
                <input type="radio" name="address" value="�ӱ�">
                �ӱ�</td>
              <td width="51"> 
                <input type="radio" name="address" value="�㶫">
                �㶫 </td>
            </tr>
            <tr> 
              <td width="53" height="19"> 
                <input type="radio" name="address" value="����">
                ���� </td>
              <td width="70" height="19"> 
                <input type="radio" name="address" value="����">
                ����</td>
              <td width="51" height="19"> 
                <input type="radio" name="address" value="�Ĵ�">
                �Ĵ� </td>
              <td width="53" height="19"> 
                <input type="radio" name="address" value="����">
                ���� </td>
              <td width="51" height="19"> 
                <input type="radio" name="address" value="ɽ��">
                ɽ�� </td>
            </tr>
            <tr> 
              <td height="16" width="53"> 
                <input type="radio" name="address" value="�Ϻ�">
                �Ϻ� </td>
              <td height="16" width="70"> 
                <input type="radio" name="address" value="ɽ��">
                ɽ��</td>
              <td height="16" width="51"> 
                <input type="radio" name="address" value="�㽭">
                �㽭 </td>
              <td height="16" width="53"> 
                <input type="radio" name="address" value="����">
                ���� </td>
              <td height="16" width="51">����*�� </td>
            </tr>
          </table>
            </td>
        </tr>
        <tr>
            
        <td width="80" height="66">����Ը���� </td>
            
        <td width="315" height="66"> 
          <textarea name="info" rows="4" cols="35"></textarea>
            ��*�� </td>
        </tr>
        <tr>
            
        <td colspan="2" bgcolor="<%=titledarkcolor%>" height="15" align=center><font color="#FFFFFF">�ͳ������˫�۳�������...</font></td>
        </tr>
        <tr>
            
        <td align="center" colspan="2" height="13">
<input
            type="submit" style="border:1 solid <%=titlelightcolor%>;color:<%=titlelightcolor%>;background-color:white" value="�͡�����">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input style="border:1 solid <%=titlelightcolor%>;color:<%=titlelightcolor%>;background-color:white" type="reset" value="�塡����"> </td>
        </tr>
    </table>
    </center></div>
</form>
</body>
</html>
<% 
end if
conn.close
set conn=nothing
set rs=nothing%>
