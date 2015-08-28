<!--#include file="conn.asp"-->
<!--#include file="connbbs.asp"-->
  <%
  menu=request("menu")
  if menu="qu" then
    call qu()
elseif menu="qu1" then
    call qu1()
elseif menu="cun" then
    call cun()
elseif menu="cun1" then
    call cun1()
	else
	call qu()
  end if
CloseDatabase		'关闭数据库
  
  %>
  
  <%sub qu() 
  Jname=session("yx8_mhjh_username")
sql="select * from 客户 where 帐号='" & Jname & "' "
set rs=conn.execute(sql)
  %>
<html>
<head>
<title>经纪人办公室</title>
</head>
<body oncontextmenu=self.event.returnValue=false; bgcolor=#FFFEF4 style="font-size:9pt;" >
<table align="center" cellspacing="1" bgcolor="#666666">
  <tr> 
    <td> 
      <p align=left style="color:#000000"><font style="font-size:9pt;" color="#FFFFFF">从股票帐户提钱</font></p>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> 
      <p align=center><font style="font-size:9pt;" color=blue>您的股票帐户里现有资金 <%=int(rs("资金"))%>两银子，您打算取多少？</font></p>
      <p align=center><%
'sqlbbs="select userWealth from [user] where username='" & session("yx8_mhjh_username") & "' "
sqlbbs="select 银两 as userWealth from [用户] where 姓名='" & session("yx8_mhjh_username") & "' "
set rsbbs=connbbs.execute(sqlbbs)
%>
      <font style="font-size:11pt;" color="#000099">你的身上有
	  <% if rsbbs("userWealth")=0 then %>
	  <%="０"%>
	  <% ELSE %>
	  <%=formatnumber(rsbbs("userWealth"),0)%>
	  <% END IF %>两银子 
	  </font> 
      <%rsbbs.close%>
</p>
      <p align=center style="color:#000000"> 
      <form method=POST action='manage.asp?menu=qu1' id=form1 name=form1  >
        <div align="center">金额： 
          <input type=text name=money  size=10>
          <input type=submit value=确定 id=submit1 name=submit1>
        </div>
      </form>
      <p align="center"><a href="stock.asp"><font style="font-size:9pt;" color="#0000FF">返回股票交易大厅</font></a></p>
      </table>
</body>
</html>
<%
conn.close
end sub
''''''''%>

<%sub qu1()
jin=request.form("money")
jin=abs(int(jin))
mess=""
Jname=session("yx8_mhjh_username")
sql="select * from 客户 where 帐号='" & Jname & "' "
set rs=conn.execute(sql)
yin=rs("资金")
if jin>yin then 
mess="你的股票经纪人大声说，我没有帮你赚这么多钱呀？"
else
sql="update 客户 set 资金=资金-"& jin & " where 帐号='" & Jname & "'"
conn.execute sql
%>

<%
'sqlbbs="update [user] set userWealth=userWealth+" & jin & " where username='" & Jname & "'"
sqlbbs="update [用户] set 银两=银两+" & jin & " where 姓名='" & Jname & "'"
connbbs.execute sqlbbs
mess="你从股票帐户取出了"&jin&" 两银子，拿好了啊"
end if
%>

<head>
<title></title>
</head>

<body oncontextmenu=self.event.returnValue=false bgcolor=#FFFEF4 style="font-size:9pt;">
<table width=100%>
  <tr>
    <td> 
      <p align=center style='color:red'><%=mess%></p>
      <p align=center><a href=stock.asp><font style="font-size:9pt;" color="#0000FF">返回</font></a></p>
    </td>
  </tr>
</table>
</body>
<%end sub''''''%>


<%sub cun()
sql="select * from 客户 where 帐号='" & session("yx8_mhjh_username") & "' "
set rs=conn.execute(sql)

%>
<html>
<head>
<title>经纪人办公室</title>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=#FFFEF4 style="font-size:9pt;">
<table align="center" cellspacing="1" bgcolor="#666666">
  <tr> 
    <td align="left" bgcolor="#666666" height="20"> 
      <p style="color:#000000"><font style="font-size:9pt;" color="#FFFFFF">存钱进股票帐户</font></p>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" align="center"> 
      <p align=center><font style="font-size:9pt;" color="blue">你的股票帐户里有资金 <font style="font-size:9pt;" color="#FF0000"><%=int(rs("资金"))%></font> 两银子，您打算存多少？ </font></p>
      <%
	  
	
'sqlbbs="select userWealth from [user] where username='"&session("yx8_mhjh_username")&"'"
sqlbbs="select 银两 as userWealth from [用户] where 姓名='"&session("yx8_mhjh_username")&"'"
set rsbbs=connbbs.execute(sqlbbs)
%>
      你的身上有<font style="font-size:9pt;" color="#000099">
	  <% if rsbbs("userWealth")=0 then %>
	  <%="０"%>
	  <% ELSE %>
	  <%=formatnumber(rsbbs("userWealth"),0)%>
	  <% END IF %>
	  </font> 两银子 
      <%
rsbbs.close

%>
      <p align=center style="color:#000000"> 
      <form method=POST action='manage.asp?menu=cun1' id=form1 name=form1 >
        金额： 
        <input type=text name=money  size=10>
        <input type=submit value=确定 id=submit1 name=submit12>
      </form>
      <p><a href="stock.asp"><font style="font-size:9pt;" color="#0000FF">返回股票交易大厅</font></a></p>
      </table>
<p align="center">&nbsp;</p>

		
</body>
</html>
<%
connbbs.close
end sub
''''''%>

<%sub cun1()
jin=request.form("money")
jin=abs(int(jin))
mess=""
Jname=session("yx8_mhjh_username")
'sqlbbs="select * from [user] where username='" & Jname & "' "
sqlbbs="select 银两 as userWealth from [用户] where 姓名='" & Jname & "' "
set rsbbs=connbbs.execute(sqlbbs)
yin=rsbbs("userWealth")
if jin>yin then 
mess="你的股票经纪人大声说，你哪里这么多钱，想坑我吗？"
else
'sqlbbs="update [user] set userWealth=userWealth-"& jin & " where username='" & Jname & "'"
sqlbbs="update [用户] set 银两=银两-"& jin & " where username='" & Jname & "'"
connbbs.execute sqlbbs
%>

<%       
sql="update 客户 set 资金=资金+" & jin & " where 帐号='" & Jname & "'"
conn.execute sql
mess="你已经把"&jin&"两银子存进了你的股票帐户"
end if
%>

<head>
<title></title>
<link rel="stylesheet" href="../css/style%5B3%5D.css" type="text/css">
</head>

<body oncontextmenu=self.event.returnValue=false bgcolor=#FFFEF4 style="font-size:9pt;">
<table width=100%>
  <tr>
    <td> 
      <p align=center style='color:red'><%=mess%></p>
      <p align=center><a href="stock.asp"><font style="font-size:9pt;" color="#0000FF">返回</font></a></p>
    </td>
  </tr>
</table>
</body>
<%end sub%>






