<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
if name="" then Response.Redirect "../../error.asp?id=016"
action=request.querystring("action")
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 物品 WHERE 所有者='" & name & "' and 数量>0 and 属性='药品' order by id "
Set Rs=conn.Execute(sql)
%>
<html><head><title>配置药材</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../style.css" type="text/css">
</head><body bgcolor="#000000" text="#FFFFFF" leftmargin="0" topmargin="0"><div align="left"> <div align="center">现已配药材一览<font face="幼圆"><a href="javascript:this.location.reload()">刷新</a></font></div><div align="center"><br> <table border="1" align="center" width="141" cellpadding="1" cellspacing="0" height="31"><tr align="center"> <td nowrap width="64"> <div align="center"><font color="#FFFFFF" size="-1">名称</font></div></td><td nowrap width="33"><div align="center"><font color="#FFFFFF" size="-1">数量 </font> </div></tr> <%
do while not rs.eof
%> <tr> <td width="64" height="3"> <div align="center"><font color="#FFFFFF" size="-1"><%=rs("名称")%>
</font> </div></td><td width="33" height="3"> <div align="center"><font color="#FFFFFF" size="-1"><%=rs("数量")%></font> </div></tr> <%
rs.movenext
loop
%> </table></div></div>
</body></html><%
rs.Close             
set rs=nothing             
conn.Close             
set conn=nothing     
%>