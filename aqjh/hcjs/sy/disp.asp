<%
dim conn,rs,id
username=Session("aqjh_name")
%>
<%if Session("aqjh_name")=""  then
  Response.Redirect "../error.asp?id=440"
else
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
id=request.querystring("id")
set rs=server.createobject("adodb.recordset")
rs.open ("SELECT * FROM sy WHERE ID=" & id),conn,0,1
%>

<html>
<head>
<title>◎  -&gt; 查看状词</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css><!--td {  font-family: 宋体; font-size: 9pt}body {  font-family: 宋体; font-size: 9pt}select {  font-family: 宋体; font-size: 9pt}A {text-decoration: none; font-family: "宋体"; font-size: 9pt}A:hover {text-decoration: underline; color: #CC0000; font-family: "宋体"; font-size: 9pt} .big {  font-family: 宋体; font-size: 12pt}.txt {  font-family: "宋体"; font-size: 10.8pt}
--></style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#3a4b91" text="#ffffff" link="#ffffff" alink="#ffffff" vlink="#ffffff" leftmargin="0" topmargin="0">
<hr size="1">   
<table width="590" border="0" cellspacing="0" cellpadding="3" align="center">
  <tr>    
    <td>受害人<b>:</b><b><%=rs("原告")%></b>&nbsp;悲惨地写道：</td>   
    <td align="right"><a href=javascript:history.back()>返 回</a></td>   
  </tr>   
</table>   
<table width="588" border="1" cellspacing="1" cellpadding="0" align="center" bordercolor="#FFFFFF" bgcolor="#000000">
  <tr bgcolor="#000000">    
    <td height="67"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td> 
            <p><span class="txt"><br>
              <%=changechr(rs("状词"))%><br>
              <br>
              <br>
              对被告这种恶劣行径，希望给于<%=rs("要求")%> <br>
              <br>
              </span></p>
            <p align="right">　</p>
          </td>
        </tr>
      </table>
    </td> 
  </tr> 
</table> 
<%rs.close   
set rs=nothing   
%>
<hr size="1">
<p align="center">　</p>
</body>   
</html>   
<%   
end if  
function getorder(theid)   
dim tmprs   
    tmprs=conn.execute("select [Order] from bbs Where ID=" & theid)   
    getorder=tmprs("Order")   
	set tmprs=nothing   
end function   
   
function changechr(str)   
    changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")   
    changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>")   
    changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>")   
end function   
%>