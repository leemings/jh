<!--#include file="conn.asp"-->
<% 
name1=trim(request.form("name"))

sql="select * from �ͻ� where �ʺ�='"&name1&"'"
set rs=conn.execute(sql)
if rs.eof then

	sql="insert into �ͻ�(�ʺ�,�ʽ�,����) values('" & name1 & "',1000,0)"
	conn.execute sql
	sql2="insert into ���ʽ�(�ʺ�) values('" & name1 & "')"
    conn.execute sql2
	mess="��Ʊ�ʻ�����ɹ���<br><font color=#000099>�������Ǹ����񣬹�Ʊ�г��㷢�ȱ�����<font color=#990000><b> 1000 </b></font> �����ӽ�������ʻ���ְɣ�</font> "

end if
conn.close
%>

<head>
<title>�����ɹ�</title>
</head>

<body bgcolor=#FFFEF4>
<table border=1 bgcolor="#666666" align=center width=350 cellpadding="10" cellspacing="13">
  <tr>
    <td bgcolor=#FFFFFF> 
      <table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><font size="3"><%=mess%></font></p>
            <p align=center><a href="stock.asp"><font color="#0000FF">����</font></a></p>
</td></tr>
</table>
    </td>
  </tr>
</table>
</body>