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
<title>轩辕设置</title>
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
	rs.open "Select * From n order by b",conn,0,1
else
	rs.open "select * from n where b='" & mp & "' order by b",conn,0,1
end if
if not rs.eof then
%><font color="#FF3333">[轩辕设置]</p>       
<p align="center">对各人轩辕武功招式修改！ <a href=adminxx_cl.asp>清理轩辕</a></p>       
<table width="680" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">       
  <tr bgcolor="#336633">        
    <td align="center" width="120"><font color=ffffff>姓 名</td>       
    <td width="90" align="center"><font color=ffffff>招式名</td>      
    <td width="38" align="center"><font color=ffffff>武功</td>      
    <td width="34" align="center"><font color=ffffff>内力</td>      
    <td width="61" align="center"><font color=ffffff>修炼次数</td>     
    <td width="150" align="center"><font color=ffffff>发招说明</td>     
    <td width="57" align="center"><font color=ffffff>招式图片</td>     
    <td width="45" align="center"><font color=ffffff>操 作</td>     
  </tr>     
  <%do while not rs.eof%>     
  <tr bgcolor="#FFFFFF"> 
    <td width="120"> 
      <div align="center"><%=rs("b")%></div> 
    </td> 
    <td width="90"> 
      <div align="center"><%=rs("a")%></div>
    </td> 
    <td width="38"> 
      <div align="center"><%=rs("c")%></div> 
    </td> 
    <td width="34"> 
      <div align="center"><%=rs("d")%></div> 
    </td>  
    <td> 
      <div align="center"><%=rs("e")%></div> 
    </td> 
    <td width="150"> 
      <div align="center"><%=rs("f")%></div>    
    </td>    
    <td width="57">     
      <div align="center"><%=rs("g")%></div>    
    </td>    
    <td width="51">     
      <div align="center">     
        <a href="delxx.asp?wgid=<%=rs("id")%>" title="删除武功">删</a> <a href="managexx.asp?wgid=<%=rs("id")%>" title="修改武功招式！">修</a></div> 
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
<div align="center">此人还未设置轩辕武功！      
<% rs.close     
set rs=nothing     
conn.close     
set conn=nothing     
end if%> </div>     
</table>     
<div align="center"></div>     
<p align="center">[<a href="adminxx.asp?mp=<%=mp%>">====刷新====</a>]     
</body>