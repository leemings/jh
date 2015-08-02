<%
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=16"
Set Cn=Server.CreateObject("ADODB.Connection")
set rst=server.createobject("adodb.recordset")
diaoyu=Application("Ba_jxqy_connstr")
Cn.Open diaoyu
rst.open"select * from 钓鱼 where 姓名='"& username &"'",cn,1,1
if rst.bof or rst.eof then
cn.execute"insert into 钓鱼(姓名,时间) values ('" & username & "',now())"
Response.Redirect "diao.asp"
else
if rst("时间")<now()-1/650 then Response.Redirect "pao.asp"
end if
%>
<html>
<head>
<title>钓鱼</title>
<meta http-equiv="refresh" content="10">
<style></style>
<link rel="stylesheet" href="../chat/READONLY/Style.css">
</head>

<BODY oncontextmenu=self.event.returnValue=false BGCOLOR="#ffffff" text="#000000">
      
<DIV ID="Layer1" STYLE="position:absolute; left:193px; top:180px; width:192px; height:165px; z-index:3"> 
<%
if rst("时间")>now()-1/720 then 
%>
<IMG SRC="diaoyu.gif" width="280" height="159">  
<%      
else      
%>       
<a href="diaoyuok.ASP"><IMG SRC="diaoyuok.gif" border="0" width="280" height="159"></a>     
<%     
end if      
%>       
</DIV>       
<DIV ID="Layer2" STYLE="position:absolute; left:112px; top:93px; width:252px; height:102px; z-index:1"><IMG SRC="diao1.jpg"></DIV>       
<div id="Layer3" style="position: absolute; left: 75; top: 333; width: 75; height: 12; z-index: 4"><font color=red>进行时间:<%=(Hour(time())*3600+Minute(time())*60+Second(time()))-(Hour(rst("时间"))*3600+Minute(rst("时间"))*60+Second(rst("时间")))%></font></div>   
<div id="Layer4" style="position:absolute; left:439px; top:51px; width:105px; height:66px; z-index:4"><img src="diao2.gif"></div>       
<table width="90%" border="1" cellspacing="0" cellpadding="0" height="90%" align="center" bgcolor="#FFFFFF">         
  <tr>          
    <td> </td>         
  </tr>         
</table>        
</body>          
</html>