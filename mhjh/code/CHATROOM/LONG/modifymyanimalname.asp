<%@ LANGUAGE=VBScript%>
<%
id=Request.QueryString ("id")
if id="" then Response.end
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
%>
<html>
<head>
<style></style>
<link rel=stylesheet href='../css.css'></head>
<body oncontextmenu=self.event.returnValue=false background='../bg1.gif' bgproperties="fixed" leftmargin="3">
<div align="center"><br>
  <font size="3"> <span style='font-size:9pt'>神 
  龙 窟</span></font><br>
<!--#include file="data.asp"--><%
rs.Open ("SELECT * FROM myanimal where rest=0 and id="&id),connb
if not(rs.EOF or rs.BOF) then
%> 
    
  <table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
    <form method="post" action="modifymyanimalnameok.asp" id=form1 name=form1>
      <tr align="center"> 
        <td colspan="2"> 原名 
          <input type="text" readonly name="animalname" size="10" maxlength="10" value="<%=rs("animalname")%>">
          <br>
          新名 
          <input type="text" name="newname" size="10" maxlength="10">
          <br>
          <input type="submit" name="Submit" value="提交">
        </td>
      </tr>
      <%
end if
rs.Close
set rs=nothing
connb.close
set connb=nothing
%> 
    </form>
  </table>
  <div align="left">你可以给你的神兽取名字,如[破破]等！<br>
    注意：请不要取一些低级下流的名字,否则站长有权去除你的神兽！</div>
</div>
<p align="center">　<input type=button value='返回' onclick="javascript:location.href='myanimal.asp'"></p>
</body>
</html>