<%@ LANGUAGE=VBScript codepage ="936" %>
<%response.expires=0
%><html>
<head>
<title>���а�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<%
const MaxPerPage=10
dim totalPut
dim CurrentPage
dim TotalPages
dim i,j
act=Request("act")
if act<>"����" and act<>"��ɱ��" and act<>"Ů����" and act<>"�и���" and act<>"�Ƹ�����" and act<>"���" and act<>"��Ա��" and act<>"�书" and act<>"����" and act<>"����" then
	act="����"
end if
%>
<script language="JavaScript">
var chatbgcolor = parent.chatbgcolor;
var chatimage = parent.chatimage;
document.write("<body oncontextmenu=self.event.returnValue=false bgcolor=" + chatbgcolor + " background=" + chatimage + " bgproperties=fixed topMargin=10><center>");
function gourl(urldz)
{
location.href="KILLMAN.ASP?act="+urldz
}
</script>
<table border="1" width="98%" cellspacing="0" cellpadding="0" bgcolor="#336633" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
    <td width="100%"> 
      <div align="center"> 
        <select name="select" onChange='gourl(this.value);'>
          <option value="����" <%if act="����" then%>selected<%end if%>>��������</option>
          <option value="��ɱ��" <%if act="��ɱ��" then%>selected<%end if%>>��ɱ����</option>
          <option value="Ů����" <%if act="Ů����" then%>selected<%end if%>>Ů �� ��</option>
          <option value="�и���" <%if act="�и���" then%>selected<%end if%>>�� �� ү</option>
          <option value="�Ƹ�����" <%if act="�Ƹ�����" then%>selected<%end if%>>�Ƹ�����</option>
          <option value="���" <%if act="���" then%>selected<%end if%>>�������</option>
          <option value="��Ա��" <%if act="��Ա��" then%>selected<%end if%>>��Ա��</option>
          <option value="�书" <%if act="�书" then%>selected<%end if%>>�书����</option>
          <option value="����" <%if act="����" then%>selected<%end if%>>��������</option>
          <option value="����" <%if act="����" then%>selected<%end if%>>��������</option>
        </select>
      </div>
</td>
</tr>
</table>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
dim sql
dim rs
dim filename
sl=20
select case act
case "��ɱ��"
	rs.open "select top "& sl &"  ����,"& act &" from [�û�] order by ��ɱ�� desc",conn,1,1
case "Ů����"
	act="���"
	rs.open "select top "& sl &"  ����,"& act &" from [�û�] where �Ա�='Ů' and grade<10 order by ��� desc",conn,1,1
case "�и���"
	act="���"
	rs.open "select top "& sl &"  ����,"& act &" from [�û�] where �Ա�='��' and grade<10 order by ��� desc",conn,1,1
case "�Ƹ�����"
	act="���"
	rs.open "select top "& sl &"  ����,"& act &" from [�û�]  where grade<10 order by ��� desc",conn,1,1
case "���"
	rs.open "select top "& sl &"  ����,"& act &" from [�û�] where grade<10 order by ��� desc",conn,1,1
case "��Ա��"
	rs.open "select top "& sl &"  ����,"& act &" from [�û�]  where grade<10 order by ��Ա�� desc",conn,1,1
case "�书"
	rs.open "select top "& sl &"  ����,"& act &" from [�û�]  where grade<10 order by �书 desc",conn,1,1
case "����"
	rs.open "select top "& sl &"  ����,"& act &" from [�û�]  where grade<10 order by ���� desc",conn,1,1
case "����"
	rs.open "select top "& sl &"  ����,"& act &" from [�û�]  where grade<10 order by ���� desc",conn,1,1
case else
	rs.open "select top "& sl &" ����,"& act &" from [�û�]  where grade<10 order by ���� ",conn,1,1
end select
if rs.eof and rs.bof then
	response.write "<p align='center'>û�п����еĶ��� </p>"
else
%>
<table border="1" cellspacing="0" width="98%" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="4" align="center">
  <tr bgcolor="#336633">
    <td align="center"><font color="#FFFFFF">�� ��</font></td>
    <td align="center"><font color="#FFFFFF"><%=act%></font></td>
</tr>
<%filename=0
do while not rs.eof%>
<tr bgcolor="#CCCCCC" onmouseover="this.bgColor='#85C2E0';" onmouseout="this.bgColor='#CCCCCC';">
    <td align="center"><%=rs("����")%></td>
    <td align="center"><%=rs(act)%></td>
</tr>
<%
rs.movenext
filename=filename+1
if filename>sl then Exit Do
loop
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</body>
</html>