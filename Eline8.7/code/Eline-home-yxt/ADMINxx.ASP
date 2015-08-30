<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
mp=Request.QueryString("mp")
%>
<html>
<head>
<title>轩辕设置♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css><!----></style>
<script language="JavaScript">
function shutwin()
{
window.close();
return;
}
</script>
</head>
<body bgcolor="#000000" text="#0000FF" vlink="#0000FF" alink="#0000FF" leftmargin="0" topmargin="0" background="../jhimg/bk_hc3w.gif">
<p align="center"> <font size="2">
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
set rs=server.createobject("adodb.recordset")
conn.open Application("sjjh_usermdb")
if mp="" then
	rs.open "Select * From n order by b",conn,0,1
else
	rs.open "select * from n where b='" & mp & "' order by b",conn,0,1
end if
if not rs.eof then
%><font color="#FF3333" size="3" face="黑体">[ </font></font><font color="#FF3333" size="3" face="黑体">轩辕设置 <font size="3">      
 ]</font></font></p>       
<p align="center"><font color="#000000" size="2">对各人轩辕武功招式修改！</font></p>       
<table width="680" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">       
  <tr bgcolor="#336633">        
    <td align="center" width="120"><font color="#FFFFFF" size="2">姓 名</font></td>       
    <td width="90" align="center"><font size="2" color="#FFFFFF">招式名</font></td>      
    <td width="38" align="center"><font color="#FFFFFF" size="2">武功</font></td>      
    <td width="34" align="center"><font color="#FFFFFF" size="2">内力</font></td>      
    <td width="61" align="center"><font size="2" color="#FFFFFF">修炼次数</font></td>     
    <td width="150" align="center"><font size="2" color="#FFFFFF">发招说明</font></td>     
    <td width="57" align="center"><font size="2" color="#FFFFFF">招式图片</font></td>     
    <td width="45" align="center"><font color="#FFFFFF" size="2">操 作</font></td>     
  </tr>     
  <%do while not rs.eof%>     
  <tr bgcolor="#FFFFFF"> 
    <td width="120"> 
      <div align="center"><font size="2"><%=rs("b")%></font></div> 
    </td> 
    <td width="90"> 
      <div align="center"><font size="2"><%=rs("a")%></font></div>
    </td> 
    <td width="38"> 
      <div align="center"><font size="2"><%=rs("c")%></font></div> 
    </td> 
    <td width="34"> 
      <div align="center"><font size="2"><%=rs("d")%></font></div> 
    </td>  
    <td> 
      <div align="center"><font size="2"><%=rs("e")%></font></div> 
    </td> 
    <td width="150"> 
      <div align="center"><font size="2"><%=rs("f")%></font></div>    
    </td>    
    <td width="57">     
      <div align="center"><font size="2"><%=rs("g")%></font></div>    
    </td>    
    <td width="51">     
      <div align="center"><font size="2">     
        <a href="delxx.asp?wgid=<%=rs("id")%>" title="删除武功">删</a> </font><font size="2"><a href="managexx.asp?wgid=<%=rs("id")%>" title="修改武功招式！">修</a></font></div> 
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
<div align="center"><font size="2" color="#000000">此人还未设置轩辕武功！</font>      
<% rs.close     
set rs=nothing     
conn.close     
set conn=nothing     
end if%> </div>     
</table>     
<div align="center"></div>     
<p align="center"><font size="2" color="#FF0000">[<a href="adminxx.asp?mp=<%=mp%>">====刷新====</a>]</font>     
</body>     
