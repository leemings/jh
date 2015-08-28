<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
lastip=Request.ServerVariables("QUERY_STRING")
if instr(lastip,"or") then response.end
my=session("yx8_mhjh_username")
if my="" then response.redirect "../error.asp?id=016"
room=request("room")
%>
<!--#include file="data.asp"-->
<%
sql="select 性别 from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
sex=rs("性别")
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
huzu=rs("户主")
rs.close
set rs=nothing
%>
<html>
<head>
<title>我 的 小 屋</title>
<style type="text/css">
<!--
td, body {  font-family: "宋体"; font-size: 12px}
a:link {  font-family: "宋体"; font-size: 12px; text-decoration: none}
a:visited {  font-family: "宋体"; font-size: 12px; text-decoration: none}
-->
</style>
</head>
<body bgcolor=#990099 text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF">
<div align="center">欢迎<%=my%>回到自己的小窝 <br>
  <br>
</div>
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
    <td align="center" width=10%><a href="../gr/pet.asp"><%if lx="豪华别墅" then%><a target="_blank" href="../gr/pet.asp"><font color="#000000">宠物</a></font><a target="_blank" href="../gr/pet.asp"><%end if%></a></td>
    <td align="center" width=10%><a href="xiaowu1.asp?room=j"><font color="#000000"><%if lx="花园洋房" or lx="豪华别墅" then%>书房</font></a><%end if%></td>
  </tr>
</table>
<br>
<table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr align="center"> 
    <td> <%if sex="女" then%><img src="image/logonan.jpg" width="139" height="269">&nbsp;&nbsp;<%end if%> 
      <%if sex="男" then%><img src="image/logonv.jpg" width="135" height="268">&nbsp;&nbsp;<%end if%>    
      <%if lx="一般房屋" then%><%if room=empty then%><img src="image/kt-01.jpg"><%else%><img src="image/ws-01.jpg"><%end if%><%end if%><%if lx="高级公寓" then%><br>   
      <%if room=empty then%><img src="image/kt-02.jpg"><%else%><img src="image/ws-02.jpg"><%end if%><%end if%>   
     <%if lx="花园洋房" then%><%if room=empty then%><img src="image/kt-03.jpg"><%else%><img src="image/ws-03.jpg"><%end if%><%end if%>   
     <%if lx="豪华别墅" then%><%if room=empty then%><img src="image/kt-04.jpg"><%else%><img src="image/ws-04.jpg"><%end if%><%end if%>    
    </td>   
  </tr>   
<tr align="center">    
    <td height="107"> 你回来了 当前位置：客厅&nbsp;   
      <% if room="b" then %>    
      <form method=POST action='pub1.asp'>   
你想要休息：   
<select name=date size=1>   
<option value="0">零天   
<option value="1">一天   
<option value="2">两天   
<option value="3">三天   
</select>   
<select name=time size=1>    
<option value="0">0小时  
<option value="1">1小时  
<option value="2">2小时  
<option value="3">3小时  
<option value="4">4小时  
<option value="5">5小时  
<option value="6">6小时  
<option value="7">7小时  
<option value="8">8小时  
<option value="9">9小时  
<option value="10">10小时  
<option value="11">11小时  
<option value="12">12小时  
<option value="13">13小时  
<option value="14">14小时  
<option value="15">15小时  
<option value="16">16小时  
<option value="17">17小时  
<option value="18">18小时  
<option value="19">19小时  
<option value="20">20小时  
<option value="21">21小时  
<option value="22">22小时  
<option value="23">23小时  
</select>  
<INPUT TYPE="submit" VALUE="确定"><INPUT TYPE="reset" VALUE="重写">   
<br>   
<br>   
体力和内力各50/小时 <br> 
当前位置：卧室&nbsp;<br>
</form>
<%
end if
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="data.asp"-->
<%
if room="c" then 
sql="SELECT * FROM 物品 where 属性='药品' and  所有者='" & my & "'"
set rs=conn.Execute(sql)
%>
<br><br>
      随身包裹   
      下面是房屋中的物品可以可以吃掉它们<br>   
      <br>   
      <table border="1" bordercolor="#000000" cellspacing="0" cellpadding="0">   
        <tr align=center>   
          <td height="20" width="200">物品名</td>   
          <td height="20" width="159">数量</td>   
          <td height="20" width="159">操作</td>   
</tr>   
<%    
do while not rs.eof   
%>   
<tr align=center>   
          <td height="20" width="200"><%=rs("名称")%></td>   
          <td height="20" width="159"><%=rs("数量")%></td>   
          <td height="20" width="159"><a href="shiyong.asp?id=<%=rs("id")%>">服用</a></td>              
</tr>  
<%   
rs.movenext   
loop  
%>  
</table>  
<br>  
<br>  
<%  
end if  
set rs=nothing  
conn.close  
set conn=nothing  
%>  
</td>  
</tr>  
</table>          
</body>  
</html>