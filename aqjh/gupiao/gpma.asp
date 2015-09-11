<!--#include file="conn.asp"-->
<!--#include file="chkmaster.asp"-->

<% 
if session("aqjh_name")="" then
response.redirect "../login.asp"
else
if not master then
call endinfo("您没有资格操作")
else %>
<html>
<head>
<title>股市管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" style="font-size:9pt;">
<table width="98%" border="0" cellspacing="1" cellpadding="3" align="center" bgcolor="#666666">
  <tr bgcolor="#000066"> 
    <td colspan="6"> 
      <table  cellspacing=0 cellpadding=1 width="97%" 
            align=center border=0>
        <tbody> 
        <tr > 
          <td align=center><a href="stock.asp"><font color=white>股票首页</font></a><font color=white>----------------------<a href="gpma.asp"><font color=white>股票管理</font></a>----------------------</font><a href="userma.asp"><font color=white>客户管理</font></a></td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
   <tr  bgcolor=#CC99FF> 
   <form name="form1" method="post" action=gpma.asp>
    <td colspan="6" align=center height=18> 
        <input type="text"  name=gname ><input type="submit" name="Submit" value="查询股票">
      </td>
       </form>
  </tr>
  <tr align="center" > 
    <td > 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;"  color=#ffffff>代号</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td > 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;" color=#ffffff>股票名称</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td > 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;" color=#ffffff>股票数量</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td > 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;" color=#ffffff>昨日收盘价</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td > 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;" color=#ffffff>目前成交价</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td width="23%"> 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;" color=#ffffff>操作</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  <% if request("gname")="" then          
sql= "select * from 股票"         
set rs=conn.execute(sql)         
DO UNTIL RS.EOF  %>
  <tr align="center"> 
    <td bgcolor="#FFFFFF" ><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font style="font-size:9pt;" color="#000066"> 
      <%=rs("sid")%> </font></a></span></td>
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><%=rs("企业")%></font></td>
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p"><%=formatnumber(rs("流通股票"),0)%> </span></font></td>
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p"><%=formatcurrency(rs("开盘价格"),3)%> </span></font></td>
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p"><%=formatcurrency(rs("当前价格"),3)%> </span></font></td>
    <td bgcolor="#FFFFFF" width="23%"> 
      <div align="center"><font style="font-size:9pt;" color="#000066"><a href="gpma.asp?gname=<%=rs("企业")%>">修改</a></font></div>
    </td>
  </tr>
  <%         
  rs.MoveNext         
  Loop         
else
  if request("action")="save" then
sid=Request("msid")
gname=Request("gname")
gtot=Request.Form ("gtot")
yp=Request.Form ("yp")
np=Request.Form ("np")
  sql= "select * from 股票 where sid="&sid        
 set rs=server.createobject("adodb.recordset")
 rs.open sql,conn,1,3
 rs("企业")=gname
 rs("流通股票")=gtot
 rs("开盘价格")=yp
 rs("当前价格")=np
 rs.update
 rs.close
 call endinfo("股票信息已经修改完毕!")
 
  else
sql="select * from 股票 where 企业 like '%"&request("gname")&"%'"
set rs=conn.execute(sql)         
   if not rs.eof then
    do while not rs.eof %>
  <tr align="center"> 
  <form method=post  action=gpma.asp?msid=<%=rs("sid")%>&action=save>
    <td bgcolor="#FFFFFF" ><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font style="font-size:9pt;" color="#000066"> 
      <%=rs("sid")%> </font></a></span></td>
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><input type=text value="<%=rs("企业")%>" size=12 name=gname></font></td>
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p">
	<input  size=16 type=text value="<%=rs("流通股票")%>" name="gtot"> </span></font></td>
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p"> 
	<input type="text" size=16 name="yp" value=<%=rs("开盘价格")%>>
 </span></font></td>
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p"> 
	<input type="text" size=16 name="np" value=<%=rs("当前价格")%>>
</span></font></td>
    <td bgcolor="#FFFFFF" width="23%"> 
      <div align="center"><font style="font-size:9pt;" color="#000066"><div align="center"><input type=submit name=submit value=执行修改></font></div>
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
end if%>
</table>
<form name="form1" method="post" action="addgp.asp">
  <table width="80%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#666666">
    <tr> 
      <td><font style="font-size:9pt;" color=#ffffff>股票名称</font></td>               
      <td><font style="font-size:9pt;" color=#ffffff>股票数量</font></td>                
      <td><font style="font-size:9pt;" color=#ffffff>昨日收盘价</font></td>                     
      <td><font style="font-size:9pt;" color=#ffffff>目前成交价</font></td>     
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF"><input type="text" size=18 name="gname"></td>
      <td bgcolor="#FFFFFF"><input type="text"  size=18 name="gtot"></td>
      <td bgcolor="#FFFFFF"><input type="text" size=18 name="yp"></td>
      <td bgcolor="#FFFFFF"><input type="text" size=18 name="np"></td>
    </tr>
	<tr>
	<td align=center colspan=4><input type="submit" name="Submit" value="增加">
    <input type="reset" name="Submit2" value="重来"></td></tr>
  </table>  
 </form>
</body>
</html>
<% 
end if'2
end if'1
CloseDatabase		'关闭数据库


sub endinfo(message) 
'-------------------------------信息提示-------------------------------
%>
<table width="<%=TableWidth%>" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="<%=Tablebackcolor%>"><tr bgcolor="<%=Tabletitlecolor%>" ><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><b>信息提示</b></td></tr><tr><td align=center height=70 bgcolor="<%=aTabletitlecolor%>"><%=message%><br></td></tr><tr><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><a href="gpma.asp">返回</a></td></tr></table>
<%end sub
%>

