<!-- #include file="setup.asp" -->
<%
if adminpassword<>session("pass") then response.redirect "admin.asp?menu=login"
log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")

id=int(Request("id"))

response.write "<center>"

select case Request("menu")

case "banner"
showbanner

case "bannerok"
rs.Open "clubconfig",Conn,1,3
rs("banner")=Request("banner")
rs("adcode")=Request("adcode")
rs.update
rs.close
%>
更新成功<br><br><a href=javascript:history.back()>返 回</a>
<%
end select





sub showbanner
%><form method="post" action="?menu=bannerok">
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>广告代码设置</td>
  </tr>
    <tr>
    <td class=a3 align=middle width="20%">顶部广告代码<br>
	<font color="#FF0000">支持HTML</font>
 </td>
    <td class=a3>
<textarea name="banner" rows="6" style="width:100%"><%=banner%></textarea></td>
  </tr>
      <tr>
    <td class=a3 align=middle width="20%">底部版权信息<br>
	<font color="#FF0000">支持HTML</font>
 </td>
    <td class=a3>
<textarea name="adcode" rows="6" style="width:100%"><%=adcode%></textarea></td>
  </tr>
  
  </table>
<input type="submit" value=" 更 新 ">
</form>
<%
end sub




htmlend

%>