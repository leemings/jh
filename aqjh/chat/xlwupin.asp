<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>��������</title>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}</script>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed" leftmargin="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<p style="font-size:9pt; color:red" align="center"> <font color="#FFFFFF"><b>�� 
  �� �� ��</b><br>
  </font><font color="#FFFFFF"><span style='font-size:9pt'><font color="#FF0000">��</font><%=aqjh_name%><font color="#FF3333">��<br>
  </font></font></span></font></p>
<div align="center"><font color="#ffffff">����$��Ǯ|����|���|��Ʒ��|����|˵��</font></div>
<table border="1" cellspacing="0" cellpadding="0" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC" align="center" >
  <tr> 
    <td width="80" > 
      <div align="center">�� Ʒ ��</div>
    </td>
    <td width="30" > 
      <div align="center">����</div>
    </td>
    <td width="30" > 
      <div align="center">����</div>
    </td>
  </tr>
  <tr bgcolor="#85C2E0" onMouseOut="this.bgColor='#85C2E0';"onMouseOver="this.bgColor='#DFEFFF';"> 
  </tr>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w1,w2,w4,w5,w6,w7,w8 from �û� where ����='"&aqjh_name&"'",conn
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>ҩ Ʒ ��</font></b></div></td>"
if not(isnull(rs("w1"))) then
	call show(rs("w1"),"ʹ��","w1")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>�� ҩ ��</font></b></div></td>"
if not(isnull(rs("w2"))) then
	call show(rs("w2"),"�¶�","w2")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>�� �� ��</font></b></div></td>"
if not(isnull(rs("w4"))) then
	call show(rs("w4"),"Ͷ��","w4")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>�� Ƭ ��</font></b></div></td>"
if not(isnull(rs("w5"))) then
	call show(rs("w5"),"�ÿ�","w5")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>�� Ʒ ��</font></b></div></td>"
if not(isnull(rs("w6"))) then
	call show(rs("w6"),"��Ʒ","w6")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>�� �� ��</font></b></div></td>"
if not(isnull(rs("w7"))) then
	call show(rs("w7"),"�ͻ�","w7")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>�� ҩ ��</font></b></div></td>"
if not(isnull(rs("w8"))) then
	call show(rs("w8"),"��ҩ","w8")
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
sub show(data,lx,lb)
data1=split(data,";")
data2=UBound(data1)
for y=0 to data2-1
	data3=split(data1(y),"|")
%>
  <tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="80"> 
      <div align="center"><%=data3(0)%> </div>
    </td>
    <td width="30"> 
      <div align="center"><%=data3(1)%> </div>
    </td>
    <td width="30"> 
	    	<div align="center"> <a href="javascript:s('/����$ 1000|����|<%=lb%>|<%=data3(0)%>|1|˵����')"><font color="#FF0000">����</font></a> 
	 </div>
    </td>
  </tr>
  <%
erase data3
next
erase data1
end sub%>
</table>
</body>
</html>