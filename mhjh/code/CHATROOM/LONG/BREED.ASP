<%@ LANGUAGE=VBScript%>
<%
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
id=Request.QueryString ("id")
if id="" then Response.end
%>
<html>
<head>
<link rel=stylesheet href='../css.css'></head>
<body oncontextmenu=self.event.returnValue=false background='../bg1.gif' bgproperties="fixed" leftmargin="3">
<div align="center"><br>
  <font size="3"> <span style='font-size:9pt'>神 
  龙 窟</span></font><br>
<!--#include file="data.asp"--><% 
rs.Open ("SELECT * FROM myanimal where rest=0 and id="&id),connb
if not(rs.EOF or rs.BOF) then
animalname=rs("animalname")
%>
</div>
<table cellpadding="0" cellspacing="0" border="0" align="center" width="100%">
    <form method="post" action="modifymyanimalnameok.asp" id=form1 name=form1 target=ps>
      <tr align="center"> 
        
      <td colspan="9"> 
        <input type="text" readonly name="animalname" size="10" maxlength="10" value="<%=animalname%>">
        <br>
         目前攻击<%=rs("attack")
%> <br>
        最大攻击<%=rs("allattack")
%><br><br>
      </td>
    </tr>
             <tr align="center" bgcolor="#FFCC33"> 
        
      <td colspan="7" bgcolor="#669900"><font color="#FFFFFF">龙 粮</font></td> 
      <td bgcolor="#669900"><font color="#FFFFFF">数 量</font></td> 
      <td bgcolor="#669900"><font color="#FFFFFF">喂养</font></td>
      </tr> <%
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open ("select * from 物品 where 所有者='"&username&"' and 属性='龙粮' and 数量>0 "),conn
do while not(rst.BOF or rst.EOF )
%> 

      <tr align="center"> 
        <td colspan="7"><%=rst("名称")%></td>
        <td> <%=rst("数量")%></td>
        <td><a href=breedok.asp?id=<%=rst("id")%>&animalname=<%=animalname%>>喂养</a></td>
        <%
rst.MoveNext
loop
%> </tr>
<%
end if
rs.Close
set rs=nothing
connb.close
set connb=nothing
rst.Close
set rst=nothing
conn.close
set conn=nothing
%> 
    </form>
  </table>
<p align="center">　<input type=button value='返回' onclick="javascript:location.href='myanimal.asp'"></p>
</body>
</html>