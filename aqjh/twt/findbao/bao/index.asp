<%
if Session("aqjh_Name")="" then
%>
<script language=vbscript>
MsgBox "�㲻�ǻ�û�е�½���߳�ʱ����ѽ��"
location.href = "javascript:self.close()"
</script>
<%end if%>
<%
set conn=server.createobject("adodb.connection") 
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
name=Session("aqjh_Name")
sql="select * from �û� where ����='" & name & "'"
set rs=conn.execute(sql)
if rs("����")<88888 then
%>
<script language=vbscript>
MsgBox "�ҵ�����88888��������û��ѽ��"
location.href = "javascript:self.close()"
</script>
<% end if%>
<%
sql="update �û� set ����=����-88888 where ����='" & name & "'"
conn.execute sql
%>
<html>
<head>
<title>���������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../css.css">
<style type="text/css">
BODY {
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
</style>
</head>

<body  background='../../bg.gif' oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5">
<div align="center"><font color="#FFFFFF"> </font> <font size="2"> <font color="#800000">[�㻨��88888�������˰��齭������������������и��ֶ�����ʿ�Ӳ�ͬ�ĵط���̽���Ĳر���Ϣ]</font><br>
</font><font size="2"><br>
  </font> <font color="#008080" size="4"><b>�� �� ð ��</b></font><font size="2"> 
  </font> </div>
<div align="center">�� </div>
<div align="center">
<table width="600" border="1" cellspacing="0" cellpadding="3" align="center" bordercolor="#000000" height="26" bordercolordark="#FFFFFF">
<tr bgcolor="#0066CC">
<td height="10" colspan="4" bgcolor="#B88230">
<div align="center"><span style="letter-spacing: 1"><font color="#FFFFFF">���±�����Ϣ</font></span></div>
</td>
</tr>
<tr>
<td width="165" height="9" bgcolor="#008080">
<div align="center"><span style="letter-spacing: 1"><font color="#FFFF00">������</font></span>
</div>
</td>
<td colspan="3" height="9" width="429" bgcolor="#008080">
<div align="center"><font color="#FFFF00">���Ӳ���</font></div>
</td>
</tr>
<%
sql="SELECT * FROM ���� order by id desc"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr>
<td width="165" height="19" bgcolor="#008080">
<div align="center"><font color="#FFFF00"><%=rs("������")%></font></div>
</td>
<td colspan="3" height="19" width="429" bgcolor="#008080">
<p align="center"><span style="letter-spacing: 1"><font color="#FFFF00">������+<%=rs("������")%>
������+<%=rs("������")%> </font></span></p>    
</td>  
</tr>  
<%  
rs.movenext  
loop  
%>  
</table>  
<table width="600" border="1" cellspacing="0" cellpadding="3" align="center" height="26" bordercolor="#000000" bordercolordark="#FFFFFF">  
<tr>  
<td height="10" bgcolor="#B88230" colspan="4" valign="middle">  
<div align="center"><font size="2" color="#FFFFFF">���±��﷢����</font></div>  
</td>  
</tr>  
 
<tr>  
<td width="126" height="9" bgcolor="#008080">  
<div align="center"><span style="letter-spacing: 1"><font color="#FFFF00">������</font></span>  
</div>  
</td>  
<td width="218" height="9" bgcolor="#008080">  
<div align="center"><font color="#FFFF00">���Ӳ���</font></div>  
</td>  
<td width="115" height="9" bgcolor="#008080">  
<div align="center"><span style="letter-spacing: 1"><font color="#FFFF00">������</font></span></div>  
</td>  
<td width="131" height="9" bgcolor="#008080">  
<div align="center"><span style="letter-spacing: 1"><font color="#FFFF00">����ʱ��</font></span></div>  
</td>  
</tr>  
<%  
sql="SELECT * FROM ���� where ����='1' order by id desc"  
Set Rs=conn.Execute(sql)  
do while not rs.bof and not rs.eof  
%> 
<tr>  
<td width="126" height="19" bgcolor="#008080">  
<div align="center"><font color="#FFFF00"><%=rs("������")%></font></div> 
</td> 
<td width="218" height="19" bgcolor="#008080"> 
<div align="center"><span style="letter-spacing: 1"><font color="#FFFF00">������+<%=rs("������")%> 
������+<%=rs("������")%> </font> </span></div>    
</td>  
<td width="115" height="19" bgcolor="#008080">  
<div align="center"><span style="letter-spacing: 1"><%=rs("����")%></span></div>  
</td>  
<td width="131" height="19" bgcolor="#008080">  
<p align="center"><span style="letter-spacing: 1"><%=rs("ʱ��")%></span></p>  
</td>  
</tr>  
<%  
rs.movenext  
loop  
conn.close  
%>  
</table>  
<br>  
<br> 
[ <b><a href="../index.asp"><font color="#FF0000">�� ½ �� ������ ʼ Ѱ ��</font></a></b> <font color="#FF0000">  
</font> ]<br>  
  <table width="600" border="1" cellspacing="1" cellpadding="0" bordercolor="#000000">
    <tr>  
<td height="94" bgcolor="#008080">  
<p style="margin-left: 5; margin-right: 5"><br>  
<font color="#FFFF00"> 
&lt;&lt;Ѱ���ض�&gt;&gt;</font> 
</p> 
        <p style="margin-left: 5; margin-right: 5; margin-bottom: 8"><font color="#FFFFFF">1����˵���ﶼ����������ŷ�ɥ����</font><font color="#FFFF00">[����µ�]</font><font color="#FFFFFF">������Ѱ�ҵ��Ļ������˵��΢����΢��������ʱ�п��ܱ���û�õ��־��������ڹµ��ˡ�</font><br> 
<br> 
<font color="#FFFFFF"> 
2�����ﶼ�ǽ�����ϡ���䱦�������ڰ���������õ��ֺǣ����������˭�õ�����Ļ���Ҫ���ܺúǣ���Ϊ���ܿ���Ϊ���ɵ�</font><font color="#FFFF00">��������</font><font color="#FFFFFF">�ǡ�</font> 
</p></td></tr></table><br></div></body></html>