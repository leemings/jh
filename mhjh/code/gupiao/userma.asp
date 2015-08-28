<!--#include file="conn.asp"-->
<!--#include file="chkmaster.asp"-->

<% 
if session("yx8_mhjh_username")="" then
response.redirect "../login.asp"
else
if session("yx8_mhjh_username")<>Application("yx8_mhjh_admin") then
call endinfo("您没有资格操作")
else %>
<html>
<head>
<title>客户管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" style="font-size:9pt;">
<table width="98%" border="0" cellspacing="1" cellpadding="3" align="center" bgcolor="#666666">
  <tr bgcolor="#000066"> 
    <td colspan="4"> 
      <table  cellspacing=0 cellpadding=1 
            align=center border=0>
        <tbody> 
        <tr > 
          <td nowrap align=center><font style="font-size:9pt;" color=#ffffff><a href="stock.asp"><font color=white>股票首页</font></a>  ----------------------<a href="userma.asp"><font color=white>客户管理</font></a>----------------------</font><a href="gpma.asp"><font color=white>股票管理</font></a></td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  <tr  bgcolor="#0099FF"> 
   <form name="form1" method="post" action=userma.asp>
    <td colspan="4" align=center height=18> 
        <input type="text"  name=userma><input type="submit" name="Submit" value="查询客户">
      </td>
       </form>
  </tr>
  <tr align="center" > 
    
    <td > 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;" color=#ffffff>客户</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td > 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;" color=#ffffff>资金</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td > 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;" color=#ffffff>状态</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
	<td > 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;" color=#ffffff>管理</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
     </tr>
	 
  <%      if request("userma")=""  then     
sql= "select * from 客户 order by 资金 desc"         
set rs=conn.execute(sql)         
DO UNTIL RS.EOF  %>
  <tr> 
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><%=rs("帐号")%></font></td>
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p"><%=rs("资金")%></span></font></td>
    <td align=center bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p"><%if rs("锁定")=1 then%>锁定<%else%>未锁<%end if%> </span></font></td>
	<td bgcolor="#FFFFFF"> <div align="center"><font style="font-size:9pt;" color="#000066"><a href="userma.asp?userma=<%=rs("帐号")%>">修改</a></font></div>
    </td>
  </tr>
  <%         
  rs.MoveNext         
  Loop         
else
  if request("action")="save" then
  sql= "select * from 客户 where 帐号='"&request("userma")&"'"         
 set rs=server.createobject("adodb.recordset")
 rs.open sql,conn,1,3
rs("资金")=request.form("money")
rs("锁定")=trim(request.form("lock"))
rs.update
rs.close
call endinfo("用户信息已经修改完毕!")
else
sql= "select * from 客户 where 帐号 like '%"&request("userma")&"%'"         
set rs=conn.execute(sql)         
 if not rs.eof then
    do while not rs.eof
 %>
  <tr> 
  <form method=post  action=userma.asp?userma=<%=rs("帐号")%>&action=save>
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><%=rs("帐号")%></font></td>
    <td align=center bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p"><input type=text name=money value="<%=rs("资金")%>"></span></font></td>
    <td align=center bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p"><select name="lock" size=1 >
<option value="<%=rs("锁定")%>" selected><%if rs("锁定")=1 then%>锁定<%else%>未锁<%end if%></option>
 
<option value="<%=1-rs("锁定")%>"><%if rs("锁定")=1 then%>不锁<%else%>锁定<%end if%></option>

 	</select> </span></font></td>
	<td bgcolor="#FFFFFF"> <div align="center"><input type=submit name=submit value=执行修改></font></div>
    </td>
	</form>
  </tr>
  <%rs.movenext
  loop
  rs.close
  else
  call endinfo("您所查找的用户不存在!")
  end if
  end if     
  end if 
%>
</table>
 </body>
</html>
<% 
end if'2
end if'1
CloseDatabase		'关闭数据库

sub endinfo(message) 
'-------------------------------信息提示-------------------------------
%>
<table align=center cellspacing=1 cellpadding="3" width="97%">
	<tr>
		<td align="center"><b>信息提示</b></td>
	</tr>
	<tr>
		<td align=center height=70 bgcolor="#CCCCCC"><%=message%><br></td>
	</tr>
	<tr>
		<td align=center height=26 bgcolor="#0099FF"><a href="userma.asp">返回</a></td>
	</tr>
</table>
<%end sub
%>

