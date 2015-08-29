<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<!--#include file="data.asp"-->
<%
sql="select * from 淘金者 where 拥有者='" & sjjh_name & "' and (金>0 or 银>0 or 铜>0)"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "你还没有矿呢!"
  window.close()
</script>
<%
response.end
end if
 jin=rs("金")*10
 yin=rs("银")*6
 tong=rs("铜")*3
 money=jin+yin+tong
 conn.execute("update 淘金者 set 金=0,银=0,铜=0 where 拥有者='" & sjjh_name & "'")
 set rs=nothing
 conn.close
 set conn=nothing
%>
<!--#include file="conn.asp"-->
<%
 conn.execute("update 用户 set 银两=银两+" & money & " where 姓名='"&sjjh_name&"'")
 conn.close
 set conn=nothing
%>
<html>
<head>
<title>矿石收购场</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 12px}
-->
</style>
</head>

<body bgcolor="#376D95">
<br>
<table width="72%" border="0" cellspacing="1" cellpadding="2" bgcolor="#999999" align="center">
  <tr align="center" bgcolor="#376D95"> 
    <td colspan="2" height="22"><font color="#FFCCCC"><b><%=sjjh_name%></b></font><b><font color="#FFFFFF">你把身上矿石全卖了，获得银两</font><font color="#FFCCCC"><%=money%></font><font color="#FFFFFF">两</font></b></td>
  </tr>
  <tr align="center" bgcolor="#376D95">
    <td colspan="2" height="26">
      <input type="submit" name="Submit" value="关闭" onclick="window.close()">
    </td>
  </tr>
</table>
</body>
</html>