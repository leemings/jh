<%@ LANGUAGE=VBScript codepage ="936" %>

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Session("aqjh_inthechat")<>"1" then %>
<script language="vbscript">
MsgBox "������������ٽ��г����ʻ���������"
window.close()
</script>
<%response.end
end if%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>�ʻ��չ���</title>
<link rel="stylesheet" href="../../css.css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="7">
<div align="center">
<table width="778" border="0" cellspacing="0" cellpadding="4">
<tr>
<td width="101" align="center"><img src="shouhua.gif" width="101" height="304"></td>
<td colspan="2" valign="top">
<p align="center">[ <font color="#FF3300">�� �� �� ��</font> ]</p> 
 
<% 
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("aqjh_usermdb") 
rs.open "SELECT w7 FROM �û� WHERE ����='" & aqjh_name& "'",conn 
 
if IsNull(rs("w7"))  then 
	rs.close 
	set rs=nothing 
	conn.close 
	set conn=nothing 
	 
 
%> 
	�ϰ循�е��� 
<table width="450" align="center" cellspacing="2" border="1" 
cellpadding="5" bordercolor="#000000"> 
<tr> 
<td><font color="#FF6600">����</font>����:-O<br> 
<%=name%>��û��Ϣ�����ֻ���һ�仨Ҳû�У���������ʲô�����͵�ɣ���</td> 
</tr> 
</table> 
<% 
else 
%><%=name%> ��Ŀǰӵ�е��ʻ��б� 
<table border="0" cellspacing="0" cellpadding="0" width="90%" align="center"> 
<tr > 
<td style="border-bottom: 1px solid rgb(250,50,50)"> 
	<div align="center"><font color="#FF6600">�ʻ�����</font></div> 
</td> 
<td style="border-bottom: 1px solid rgb(250,50,50)"> 
	<div align="center"><font color="#FF6600">��  ��</font></div> 
</td> 
<td style="border-bottom: 1px solid rgb(250,50,50)"> 
	<div align="right"><font color="#FF6600">��  ��</font></div> 
</td> 
<td style="border-bottom: 1px solid rgb(250,50,50)"> 
	<div align="right"><font color="#FF6600">�չ��۸�</font></div> 
</td> 
<td style="border-bottom: 1px solid rgb(250,50,50)"> 
	<div align="center"><font color="#FF6600">��  ��</font></div> 
</td> 
</tr> 
<% 
d7=rs("w7") 
 
rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
 
 
if not IsNull(d7) then 
	call show(d7,"�ʻ�") 
end if 
 
 
end if 
 
sub show(data,lx) 
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("aqjh_usermdb") 
data1=split(data,";") 
data2=UBound(data1) 
 
for y=0 to data2-1 
	data3=split(data1(y),"|") 
	rs.open "SELECT h FROM b WHERE a='" & data3(0)& "'",conn 
	if  not rs.bof and not rs.eof  then 
%> 
  <tr bgcolor="#dfefff" onmouseout="this.bgColor='#DFEFFF';"onmouseover="this.bgColor='#85C2E0';">  
    <td>  
      <div align="center"><%=data3(0)%> </div> 
    </td> 
	 <td>  
      <div align="center"><%=data3(0)%> </div> 
    </td> 
    <td>  
      <div align="right"><%=data3(1)%> </div> 
    </td> 
    <td> 
      <div align="right"><%=rs("h")%><%if data3(0)="��ѩ" or data3(0)="����" or data3(0)="��ܿ" or data3(0)="����" or data3(0)="����" then%>��<%else%>��<%end if%>/��</div> 
    </td> 
    <td> 
      <div align="right"><%if data3(0)="��ѩ" or data3(0)="����" or data3(0)="��ܿ" or data3(0)="����" or data3(0)="����" then%><%=int(rs("h")/4)%>��<%else%><%=int(rs("h")/3)%>��<%end if%>/�� 
      </div> 
    </td> 
    <td>  
	    <div align="center"> <a href="dan2.asp?wpname=<%=data3(0)%>&lx=<%=data3(0)%>" ><font color="#0000FF">����</font></a> </div> 
    </td> 
  </tr> 
  <% 
  end if 
  rs.close 
erase data3 
next 
erase data1 
set rs=nothing 
conn.close 
set conn=nothing 
end sub 
%> 
 
</table> 
</td> 
</tr> 
</table> 
<br>��Ȩ���С����齭����վ��</div>
</body>
</html>