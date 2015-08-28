<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
my=session("yx8_mhjh_username")
if my="" then response.redirect "../error.asp?id=016"
%>
<!--#include file="data.asp"-->
<%
sql="select 性别,银两 from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
sex=rs("性别")
yin=rs("银两")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="data1.asp"-->
<%
sql="select * from 房屋 where 户主='" & my & "' or 伴侣='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%>
<script language=vbscript>
  MsgBox "您还没有买房子呢。"
  location.href = "../11.htm"
</script>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
if rs("状态")<>"正常" then
%>
<script language=vbscript>
  MsgBox "您的房屋是不是有点状况呀。"
  location.href = "../jh.asp"
</script>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
lx=rs("类型")
%>
<html>
<head>
<title>我 的 小 屋</title>
<style type="text/css">
<!--
td,body,input {  font-family: "宋体"; font-size: 12px}
a:link {  font-family: "宋体"; font-size: 12px; text-decoration: none}
a:visited {  font-family: "宋体"; font-size: 12px; text-decoration: none}
-->
</style>
</head>
<body bgcolor=#FFFFFF text="#000000" link="#FFFFFF" vlink="#FFFFFF">
<div align="center"> 欢迎<%=my%>回到自己的小窝 
  <table width='95%' align=center cellspacing=0 border=1 cellpadding=2 bordercolor="#000000">
    <tr bgcolor="#C4DEFF"> 
      <td align="center" width=10%><a href="xiaowu.asp"><font color="#000000"><%if lx="一般房屋" or lx="高级公寓" or lx="花园洋房" or lx="豪华别墅" then%>客厅</font></a><%end if%></td>
      <td align="center" width=10%><a href="xiaowu.asp?room=b"><font color="#000000"><%if lx="一般房屋" or lx="高级公寓" or lx="花园洋房" or lx="豪华别墅" then%>卧室</font></a><%end if%></td>
      <td align="center" width=10%><a href="xiaowu.asp?room=c"><font color="#000000"><%if lx="一般房屋" or lx="高级公寓" or lx="花园洋房" or lx="豪华别墅" then%>储藏室</font></a><%end if%></td>
      <td align="center" width=10%><a href="xiaowu1.asp"><font color="#000000"><%if lx="高级公寓" or lx="花园洋房" or lx="豪华别墅" then%>餐厅</font></a><%end if%></td>
      <td align="center" width=10%><a href="xiaowu1.asp?room=e"><font color="#000000"><%if lx="高级公寓" or lx="花园洋房" or lx="豪华别墅" then%>卫生间</font></a><%end if%></td>
      <td align="center" width=10%><a href="xiaowu2.asp"><font color="#000000"><%if lx="高级公寓" or lx="花园洋房" or lx="豪华别墅" then%>小金库</font></a><%end if%></td>
      <td align="center" width=10% bgcolor="#C4DEFF"><a href="xiaowu1.asp?room=g"><font color="#000000"><%if lx="花园洋房" or lx="豪华别墅" then%>花园</font></a><%end if%></td>
      <td align="center" width=10%><a href="xiaowu3.asp"><font color="#000000"><%if lx="豪华别墅" then%>游泳池</font></a><%end if%></td>
      <td align="center" width=10%><a href="xiaowu1.asp?room=j"><font color="#000000"><%if lx="花园洋房" or lx="豪华别墅" then%>书房</font></a><%end if%></td>
    </tr>
  </table>
  <br>
</div>
<table width="95%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolor="#000000">
  <tr> 
    <td width="40%" rowspan="2" align="center"><%if sex="女" then%><img src="image/logonan.jpg" width="177" height="298"><%end if%><%if sex="男" then%><img src="image/logonv.jpg"><%end if%></td>
    <td height="108" valign="top" align="center">金库里有<%if sex="男" then%><%=rs("银两")%><%end if%><%if sex="女" then%><%=rs("伴侣银两")%><%end if%> 
      两 <br>
      <br>
      我要存钱在家里 
      <form method="post" action="putmoney.asp">
        <input type="text" name="money" size="8" value="10">
        <input type="submit" name="a" value="确定">
        <input type="reset" name="b" value="取消">
      </form>
    </td>
  </tr>
  <tr>
    <td align="center" height="48">我要取家里的钱 
      <form method="post" action="getmoney.asp">
        <input type="text" name="money" size="8" value="10">
        <input type="submit" name="c" value="确定">
        <input type="reset" name="d" value="取消">
      </form>
    </td>
  </tr>
</table> 
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

