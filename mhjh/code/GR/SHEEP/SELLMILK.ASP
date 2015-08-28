<!--#include file="data2.asp"-->
<%
if session("yx8_mhjh_username")="" then%>
<%else
sheepname=request.form("sheepname")
milk=request.form("milk")
if sheepname="" or not isnumeric(milk) then
%>
<script language="Vbscript">
msgbox"输入错误！",0,"警告" 
history.back
</script>
<%
else
intmilk=milk-int(milk)
if intmilk<>0 or milk<1 then
%>
<script language="Vbscript">
msgbox"输入错误！",0,"警告" 
history.back
</script>
<%else
rs.open"select * from sheep where user='"&session("yx8_mhjh_username")&"' and sheepname='"&sheepname&"'",conn,1,1
if rs.bof then
rs.close
conn.close
response.redirect"warning.htm"
else
milklast=rs("milk")-milk
if milklast<0 then
rs.close
conn.close%>
<script language="Vbscript">
msgbox"你的小孩没有打这么多的工作单位啊！",0,"警告" 
history.back
</script><%
else
rs.close
conn.execute"update sheep set milk='"&milklast&"' where user='"&session("yx8_mhjh_username")&"' and sheepname='"&sheepname&"'"
rs.open"select * from rules",conn,1,1
milkprice=rs("milkprice")
rs.close
tempsplosh=milkprice*milk
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
conn.execute"update 用户 set 体力=体力+'"&tempsplosh&"' where 姓名='"&session("yx8_mhjh_username")&"'"
conn.close
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>孤儿院</title>
<LINK href="../../style4.css" rel=stylesheet>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false>

<div align="center">
</div>
<div align="center">
  <center>
  <table border="0" width="470" bordercolor="#FFFFFF" cellspacing="1"
  cellpadding="0" height="20">
    <tr>
      <td class="p6"><font size="2">□-您当前的位置-&gt;孤儿院〉&gt;领工资&nbsp;&nbsp;&nbsp; 
        <a href="javascript:history.back(1)">返回</a></font></td>
    </tr>
  </table>
  </center>
</div>
<div align="center">
  <center>
  <table border="0" width="470" cellspacing="1" cellpadding="0" height="50">
    <tr>
      <td class="p2" width="100%">
        <p align="center"><font size="2">恭喜！ 
        您已经成功领了你的孩子的工资，共获得<%=tempsplosh%>体力</font></p>
      </td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>

<% 
end if 
end if 
end if 
end if 
end if 
%>