<!--#include file="conn.asp"-->
<% 
name1=trim(request.form("name"))

sql="select * from 客户 where 帐号='"&name1&"'"
set rs=conn.execute(sql)
if rs.eof then

	sql="insert into 客户(帐号,资金,锁定) values('" & name1 & "',1000,0)"
	conn.execute sql
	sql2="insert into 总资金(帐号) values('" & name1 & "')"
    conn.execute sql2
	mess="股票帐户申请成功！<br><font color=#000099>看到你是个菜鸟，股票市场便发慈悲送了<font color=#990000><b> 1000 </b></font> 两银子进了你的帐户里，乐吧？</font> "

end if
conn.close
%>

<head>
<title>开户成功</title>
</head>

<body bgcolor=#FFFEF4>
<table border=1 bgcolor="#666666" align=center width=350 cellpadding="10" cellspacing="13">
  <tr>
    <td bgcolor=#FFFFFF> 
      <table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><font size="3"><%=mess%></font></p>
            <p align=center><a href="stock.asp"><font color="#0000FF">返回</font></a></p>
</td></tr>
</table>
    </td>
  </tr>
</table>
</body>