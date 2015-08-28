<%@ LANGUAGE=VBScript%>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
%>

<html>
<head>
<link rel=stylesheet href='../../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../bg1.gif' bgproperties="fixed" leftmargin="1" topmargin="0">
<div align="center"><br>
<font size="3" color="#FFFFFF"><br>
  </font><b>困 龙 窟</b><br>    
<!--#include file="data.asp"--><%   
sql= "SELECT * FROM myanimal where outtime=true"    
rs.open sql,connb,1,1    
if not(rs.EOF or rs.BOF) then     
%>    
  <table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">    
<%    
do while not (rs.eof or rs.bof)    
gong=rs("attack")    
outtime=rs("outtime")    
%>     
    <tr align="center">     
      <td colspan="2"><%=rs("username")%>-<%=rs("lei")%><br>    
      </td>    
    </tr>    
      
    <tr>     
      <td colspan="2">       
<hr noshade size="1" color=darkred>    
      </td>    
    </tr>    
        
    <%       
rs.movenext    
loop    
rs.Close    
set rs=nothing    
connb.close    
set connb=nothing    
%>     
  </table>    
<table width="100%" border="0" cellspacing="2" cellpadding="2">    
<tr>    
      <td>官府可以释放神龙,龙主也可以使用银两疏通得回</td>    
</tr>    
</table>    
<br>    
  <%    
else    
%>    
  Sorry！现在没有神龙被监禁！      
  <%end if%>        
</div>        
</body>        
</html>  
