<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
if name="" then Response.Redirect "../../error.asp?id=016"
action=request.querystring("action")
name=session("yx8_mhjh_username")
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 物品 WHERE 所有者='" & name & "' and 数量>0 order by id "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then 
Response.write "你什么物品都没有"
Response.end
end if
%>
<html><head><title>药材</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../style2.css" type="text/css">
</head><body bgcolor="#000000" text="#FFFFFF" leftmargin="0" topmargin="0"><div align="left"> <div align="center">现有药材一览<font face="幼圆"><a href="javascript:this.location.reload()">刷新</a></font></div><div align="center"><br> <table border="1" align="center" width="141" cellpadding="1" cellspacing="0" height="31"><tr align="center"> <td nowrap width="64" height="16"> <div align="center"><font color="#FFFFFF">名称</font></div></td><td nowrap width="33" height="16"><div align="center"><font color="#FFFFFF">数量 </font> </div><td colspan="2" nowrap width="36" height="16"><div align="center"><font color="#FFFFFF">价格</font></div></td></tr><%
do while not rs.eof
%> <tr> <form method=POST action='wupin1.asp?action=<%=rs("名称")%>&name=<%=name%>'>
<td width="64" height="3"> <div align="center"><font color="#FFFFFF"><%=rs("名称")%></font> </div></td><td width="33" height="3"> <div align="center"><font color="#FFFFFF"><%=rs("数量")%></font> </div><td colspan="2" height="3" width="36"> <div align="center"><font color="#FFFFFF"><%=rs("价格")%></font> </div></td></form></tr> <%
rs.movenext
loop
%> </table></div></div>
</body></html>
<%
rs.Close             
set rs=nothing             
conn.Close             
set conn=nothing  
%>