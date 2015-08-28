<%
if session("yx8_mhjh_username")="" then
response.redirect "../error.asp?id=016"
response.end
end if
room=request("room")
my=session("yx8_mhjh_username")
sex=session("yx8_mhjh_usersex")
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
huzu=rs("户主")
rs.close
set rs=nothing
conn.close
set conn=nothing
if lx="普通公寓" then
%>
<script language="vbscript">
Msgbox "你的房屋没有该设备!!"
history.back(-1)
</script>
<%
response.end
end if
%>
<html>
<head>
<title>我 的 小 屋</title>
<style type="text/css">
<!--
td,body {  font-family: "宋体"; font-size: 12px}
a:link {  font-family: "宋体"; font-size: 12px; text-decoration: none}
a:visited {  font-family: "宋体"; font-size: 12px; text-decoration: none}
-->
</style>
</head>
<body bgcolor=#990099 text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF">
<div align="center">欢迎<%=my%>回到自己的小窝 <br>
  <br>
</div>
<TABLE width='95%' ALIGN=center CELLSPACING=0 BORDER=1 CELLPADDING=2 bordercolor="#000000">
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
<table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr align="center"> 
    <td> <%if sex="女" then%><img src="image/logonan.jpg" width="145" height="306">&nbsp;&nbsp;<%end if%> 
      <%if sex="男" then%><img src="image/logonv.jpg" width="141" height="306">&nbsp;&nbsp;<%end if%>     
     <%if lx="高级公寓" then%><%if room=empty then%><img src="image/ct-02.jpg"><%end if%><%if room="e" then%><img src="image/wsj-02.jpg"><%end if%><%if room="g" then%><img src="image/hy-01.jpg"><%end if%><%if room="j" then%><img src="image/sf-01.jpg"><%end if%><%end if%>    
     <%if lx="花园洋房" then%><%if room=empty then%><img src="image/ct-03.jpg"><%end if%><%if room="e" then%><img src="image/wsj-03.jpg"><%end if%><%if room="g" then%><img src="image/hy-01.jpg"><%end if%><%if room="j" then%><img src="image/sf-01.jpg"><%end if%><%end if%>    
     <%if lx="豪华别墅" then%><%if room=empty then%><img src="image/ct-04.jpg"><%end if%><%if room="e" then%><img src="image/wsj-04.jpg"><%end if%><%if room="g" then%><img src="image/hy-02.jpg"><%end if%><%if room="j" then%><img src="image/sf-02.jpg"><%end if%><%end if%>    
    </td>    
  </tr>    
<tr align="center">     
    <td height="107">     
<% if room=empty then %>     
<!--#include file="data.asp"-->     
<%    
sql="SELECT * FROM 物品 where 属性='食品' and 所有者='" & my & "'"    
Set Rs=conn.Execute(sql)    
%>     
<br>    
      增加体力、内力<br>    
      <table cellspacing="0" cellpadding="0" border="1" bordercolor="#000000">    
        <tr align="center" bgcolor="#333333">     
          <td height="20" width="159">物品名</td>    
          <td height="20" width="159">数量</td>    
          <td height="20" width="159">操作</td>    
</tr>    
<%     
do while not rs.eof and not rs.bof     
%>    
        <tr align="center">     
          <td height="20" width="159"><%=rs("名称")%></td>    
          <td height="20" width="159"><%=rs("数量")%></td>    
          <td height="20" width="159"><a href=fuyong.asp?id=<%=rs("id")%>>服用</a></td>                
</tr>    
<%     
rs.movenext     
loop    
rs.close    
set rs=nothing    
conn.close    
set conn=nothing    
%>    
</table>    
      <br>    
<%    
response.end    
end if    
%>    
<% if room="j" then%>    
<!--#include file="data.asp"-->    
<%   
sql="SELECT * FROM 物品 where 数量>0 and 属性='礼品' and 所有者='" & my & "'"   
Set Rs=conn.Execute(sql)   
%>    
<br>   
书房拆礼物 增加美貌、个性值<br>    
      <table border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">    
        <tr align="center">     
          <td height="1" width="159" bgcolor="#333333">物品名</td>    
          <td height="1" width="159" bgcolor="#333333">数量</td>    
          <td height="1" width="159" bgcolor="#333333">操作</td>    
</tr>    
<%     
do while not rs.eof and not rs.bof     
%>    
        <tr align="center">     
          <td height="1" width="159"><%=rs("名称")%></td>    
          <td height="1" width="159"><%=rs("数量")%></td>    
          <td height="1" width="159"><a href=kanshu.asp?id=<%=rs("id")%>>拆礼物</a></td>                
</tr>    
<%     
rs.movenext     
loop    
rs.close    
set rs=nothing    
conn.close    
set conn=nothing    
%>    
</table>    
<br>    
<%end if%>    
</td>    
</tr>    
</table>            
</body>    
</html>