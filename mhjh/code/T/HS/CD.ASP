<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
myname=session("yx8_mhjh_username")
if myname="" then Response.Redirect "../error.asp?id=016"
%>
<html>
<head>
<title>超渡亡魂</title>
<link rel=stylesheet href='../../style.css'></head>
<body oncontextmenu=self.event.returnValue=false background='../../chatroom/bg1.gif'>
<div align="center">
<%
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="select 姓名,门派,国家,等级 from 用户 where 形态='鬼魂' order by 等级 desc"
rs.open sql,conn,1,1
if rs.eof or  rs.bof then
response.write "江湖里一个鬼魂都没有，太和平了"
else
%>
<table border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" height="1">
<tr>
<td align="center" height="1" width="188" bgcolor="#FFFFFF">姓 名</td>
<td align="center" height="1" width="188" bgcolor="#FFFFFF">门 派</td>
<td align="center" height="1" width="188" bgcolor="#FFFFFF">国 家</td>
<td align="center" height="1" width="188" bgcolor="#FFFFFF">等 级</td>
<td align="center" height="1" width="188" bgcolor="#FFFFFF"><select name="shihe">
                            <option value="我佛慈悲" selected>选择禅功</option>                
                            <option value="我佛慈悲">我佛慈悲</option>
                            <option value="佛光普照">佛光普照</option>
                            <option value="万佛朝宗">万佛朝宗</option>
              </select></td>
</tr>
<%do while not rs.eof%>
<tr>
<td align="center" width="188" height="1"><%=rs("姓名")%>
</td>
<td align="center" width="188" height="1"><%=rs("门派")%>
</td>
<td align="center" width="188" height="1"><%=rs("国家")%>
</td>
<td align="center" width="188" height="1"><%=rs("等级")%>
</td>
<td align="center" width="188" height="1"><a href=# onclick="javascript:location.href=('cd2.asp?id=<%=rs("姓名")%>&shihe='+document.all('shihe').value);"><span class="calen-curr"><font color="#0000FF">超 度</font></span></a> 
</td> 
</tr> 
<% 
rs.movenext 
loop 
end if 
rs.Close 
set rs=nothing 
conn.Close              
set conn=nothing  
%> 
</table> 
　 
</div> 
</body> 
</html> 
