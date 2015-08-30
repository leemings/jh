<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%><!--#include file="data.asp"--><%
sql="select * from s"
set rs=conn1.execute(sql)
%>

<html>
<head>
<title>申请创派♀wWw.happyjh.com♀</title>
<link rel="stylesheet" href="../dg/Setup.css">
</head>
<BODY oncontextmenu=self.event.returnValue=false text=#ffffff background="../JHIMG/BK_HC3W01.GIF">
<div align="center">
  <table border="1" width="61%">
    <tr> 
      <td width="22%" align="center"><font color="#0000FF">申请人</font></td>
      <td width="49%" align="center"><font color="#0000FF">申请时间</font></td>
      <td width="29%" align="center"><font color="#0000FF">申请结果</font></td>
    </tr>
    
    
      <%do while not rs.bof and not rs.eof%>
      <tr> 
      <td width="22%" align="center" height="27"><font color="#0000FF"><%=rs("姓名")%></font></td>
      <td width="49%" align="center" height="27"><font color="#0000FF"><%=rs("申请时间")%></font></td>
      <td width="29%" align="center" height="27"><font color="#0000FF"> 
        <%if rs("申请结果")="未定" then
    	if sjjh_name=application("sjjh_user") then%>
        <a href="pz.asp?id=<%=rs("申请人")%>"> 批准</a> | <a href="delgf.asp?id=<%=rs("申请人")%>"> 删除</a>   
        <%else%>  
        待批  
        <%end if 
    else%>
        <%=rs("申请结果")%>   
        <%if sjjh_name=application("sjjh_user") then%>  
         | <a href="delgf.asp?id=<%=rs("申请人")%>">删除</a>  
         <%end if  
        end if%>
        </font></td></tr><%rs.movenext
  loop
  rs.close
  set rs=nothing
  conn1.close
  set conn=nothing%>
    
    
  </table>
</div>
<table width="536" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="yellow" bordercolordark="#FFFFFF" height="138">
  <tr align="center"> 
    <td bgcolor="#000080" width="526"><img border="0" src="22.GIF"></td>
  </tr>
  <tr align="center"> 
    <td height="67" bgcolor="#000080" width="526">  要自立门户首先要满足以下条件：
      <p><font color="#FF00FF">魅力</font>超过<b>100000</b>，<font color="#FF00FF">道德</font>超过<b>100000，</b><font color="#FF00FF">银两</font>超过<b>10亿，</b><font color="#FF00FF">战斗等级</font>超过<b>65级</b></p>
      <p><font color="#FF00FF">拉人数</font>超过<b>20人<br>注：所拉来的人的战斗等级在20级以上才算数</b></p>
    </td>
  </tr>
  <tr align="center"> 
    <td bgcolor="#000080" width="526">符合以上条件的用户请<a href="gfsq1.asp"><b><font color="#00FFFF">进来</font></b></a>......</td>
  </tr>
</table>
</body>
</html>
<%
%>