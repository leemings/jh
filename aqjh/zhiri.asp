<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="select c from sm where a='pk值日'"
Set Rs=conn.Execute(sql)
%>
<html>
<head>
<title>《快乐江湖》值班人员</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css.css">
</head>
<body oncontextmenu=self.event.returnValue=false leftmargin="0" topmargin="20" bgcolor="#660000">
<table width="100%" border="0" align="center">
  <tr align="center">
    <td> 
      <p align="center"><font size="2"><font color=#CCFFCC size=2>官府人员值班时间表:<br>
        <br>
        <%=rs("c")%></font></font></p>
    </td>
</tr></table>
</body>
</html>
<%set rs=nothing	
	conn.close
	set conn=nothing %>