<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
mp=Request.QueryString("mp")
%>
<html>
<head>
<title>�书����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
<script language="JavaScript">
function shutwin()
{
window.close();
return;
}
</script>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"> 
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
set rs=server.createobject("adodb.recordset")
conn.open Application("aqjh_usermdb")
if mp="" then
	rs.open "Select * From y order by b",conn,0,1
else
	rs.open "select * from y where b='" & mp & "' order by b",conn,0,1
end if
if not rs.eof then
%><font color="#FF3333">[�书����]</p>       
<p align="center"><font color="#000000" size="2">�Ը����书��ʽ�޸ģ�</font> <a href=adminwg_cl.asp><font color="#000000" size="2">�����书</a></p>       
<table width="692" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">       
  <tr bgcolor="#336633">        
    <td align="center" width="120"><font color="#FFFFFF">�� ��</font></td>       
    <td width="90" align="center"><font color="#FFFFFF">��ʽ��</font></td>      
    <td width="38" align="center"><font color="#FFFFFF">�书</font></td>      
    <td width="34" align="center"><font color="#FFFFFF">����</font></td>      
    <td width="67" align="center"><font color="#FFFFFF">ʹ�ô���</font></td>     
    <td width="150" align="center"><font color="#FFFFFF">����˵��</font></td>     
    <td width="57" align="center"><font color="#FFFFFF">��ʽͼƬ</font></td>     
    <td width="51" align="center"><font color="#FFFFFF">�� ��</font></td>     
  </tr>     
  <%do while not rs.eof%>     
  <tr bgcolor="#FFFFFF"> 
    <td width="120"> 
      <div align="center"><%=rs("b")%></font></div> 
    </td> 
    <td width="90"> 
      <div align="center"><%=rs("a")%></font></div>
    </td> 
    <td width="38"> 
      <div align="center"><%=rs("c")%></font></div> 
    </td> 
    <td width="34"> 
      <div align="center"><%=rs("d")%></font></div> 
    </td>  
    <td> 
      <div align="center"><%=rs("e")%></font></div> 
    </td> 
    <td width="150"> 
      <div align="center"><%=rs("f")%></font></div>    
    </td>    
    <td width="57">     
      <div align="center"><%=rs("g")%></font></div>    
    </td>    
    <td width="51">     
      <div align="center">     
        <a href="delwg.asp?wgid=<%=rs("id")%>" title="ɾ���书">ɾ</a> </font><a href="managewg.asp?wgid=<%=rs("id")%>" title="�޸��书��ʽ��">��</a></font></div> 
    </td>     
  </tr>     
  <%rs.movenext     
loop      
rs.close     
set rs=nothing     
conn.close     
set conn=nothing%>     
  <% else %>     
</table>     
<table align="center" width="197">     
<td height="14">     
<div align="center"><font color="#000000">�����ɻ�δ�����书��</font>      
<% rs.close     
set rs=nothing     
conn.close     
set conn=nothing     
end if%> </div>     
</table>     
<div align="center"></div>     
<p align="center"><font color="#FF0000">[<a href="adminwg.asp?mp=<%=mp%>">====ˢ��====</a>]</font>     
</body>