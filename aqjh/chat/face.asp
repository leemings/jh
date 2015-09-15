<%@ LANGUAGE=VBScript codepage ="936" %><%response.expires=0
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 名单头像 from 用户 where 姓名='" & aqjh_name &"'",conn
tx=rs("名单头像")
rs.close
set rs=nothing
conn.close
set conn=nothing
%><html>
<head>
<title>更改头像</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed" leftmargin="0" topmargin="20"><form method="POST" action="changeface.asp">
<div align="center"><font color="#FFFF00"><b>更改头像</b><br>
</font><font color="#FFFF00"><br>
头像：<img id=face src="<%=tx%>">
<select name=face size=1 onChange="document.images['face'].src='../ico/'+options[selectedIndex].value+'-2.gif';" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">
<%for i=0 to 185%>
<option value='<%=i%>'><%=i%></option>
<%next%>
</select>
<br>
<br>
<br>
<br>
<input type="submit" value="修改" name="B1" tyle="font-family: Tahoma; font-size: 12px">
</font></div>
</from>
</body>
</html>