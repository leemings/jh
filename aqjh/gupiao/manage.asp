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
CloseDatabase		'�ر����ݿ�
  
  %>
  
  <%sub qu() 
  Jname=session("aqjh_name")
sql="select * from �ͻ� where �ʺ�='" & Jname & "' "
set rs=conn.execute(sql)
  %>
<html>
<head>
<title>�����˰칫��</title>
</head>
<body oncontextmenu=self.event.returnValue=false; bgcolor=#FFFEF4 style="font-size:9pt;" >
<table align="center" cellspacing="1" bgcolor="#666666">
  <tr> 
    <td> 
      <p align=left style="color:#000000"><font style="font-size:9pt;" color="#FFFFFF">�ӹ�Ʊ�ʻ���Ǯ</font></p>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> 
      <p align=center><font style="font-size:9pt;" color=blue>���Ĺ�Ʊ�ʻ��������ʽ� <%=int(rs("�ʽ�"))%>�����ӣ�������ȡ���٣�</font></p>
      <p align=center><%
'sqlbbs="select userWealth from [user] where username='" & session("aqjh_name") & "' "
sqlbbs="select ���� as userWealth from [�û�] where ����='" & session("aqjh_name") & "' "
set rsbbs=connbbs.execute(sqlbbs)
%>
      <font style="font-size:11pt;" color="#000099">���������
	  <% if rsbbs("userWealth")=0 then %>
	  <%="��"%>
	  <% ELSE %>
	  <%=formatnumber(rsbbs("userWealth"),0)%>
	  <% END IF %>������ 
	  </font> 
      <%rsbbs.close%>
</p>
      <p align=center style="color:#000000"> 
      <form method=POST action='manage.asp?menu=qu1' id=form1 name=form1  >
        <div align="center">�� 
          <input type=text name=money  size=10>
          <input type=submit value=ȷ�� id=submit1 name=submit1>
        </div>
      </form>
      <p align="center"><a href="stock.asp"><font style="font-size:9pt;" color="#0000FF">���ع�Ʊ���״���</font></a></p>
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
Jname=session("aqjh_name")
sql="select * from �ͻ� where �ʺ�='" & Jname & "' "
set rs=conn.execute(sql)
yin=rs("�ʽ�")
if jin>yin then 
mess="��Ĺ�Ʊ�����˴���˵����û�а���׬��ô��Ǯѽ��"
else
sql="update �ͻ� set �ʽ�=�ʽ�-"& jin & " where �ʺ�='" & Jname & "'"
conn.execute sql
%>

<%
'sqlbbs="update [user] set userWealth=userWealth+" & jin & " where username='" & Jname & "'"
sqlbbs="update [�û�] set ����=����+" & jin & " where ����='" & Jname & "'"
connbbs.execute sqlbbs
mess="��ӹ�Ʊ�ʻ�ȡ����"&jin&" �����ӣ��ú��˰�"
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
      <p align=center><a href=stock.asp><font style="font-size:9pt;" color="#0000FF">����</font></a></p>
    </td>
  </tr>
</table>
</body>
<%end sub''''''%>


<%sub cun()
sql="select * from �ͻ� where �ʺ�='" & session("aqjh_name") & "' "
set rs=conn.execute(sql)

%>
<html>
<head>
<title>�����˰칫��</title>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=#FFFEF4 style="font-size:9pt;">
<table align="center" cellspacing="1" bgcolor="#666666">
  <tr> 
    <td align="left" bgcolor="#666666" height="20"> 
      <p style="color:#000000"><font style="font-size:9pt;" color="#FFFFFF">��Ǯ����Ʊ�ʻ�</font></p>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" align="center"> 
      <p align=center><font style="font-size:9pt;" color="blue">��Ĺ�Ʊ�ʻ������ʽ� <font style="font-size:9pt;" color="#FF0000"><%=int(rs("�ʽ�"))%></font> �����ӣ����������٣� </font></p>
      <%
	  
	
'sqlbbs="select userWealth from [user] where username='"&session("aqjh_name")&"'"
sqlbbs="select ���� as userWealth from [�û�] where ����='"&session("aqjh_name")&"'"
set rsbbs=connbbs.execute(sqlbbs)
%>
      ���������<font style="font-size:9pt;" color="#000099">
	  <% if rsbbs("userWealth")=0 then %>
	  <%="��"%>
	  <% ELSE %>
	  <%=formatnumber(rsbbs("userWealth"),0)%>
	  <% END IF %>
	  </font> ������ 
      <%
rsbbs.close

%>
      <p align=center style="color:#000000"> 
      <form method=POST action='manage.asp?menu=cun1' id=form1 name=form1 >
        �� 
        <input type=text name=money  size=10>
        <input type=submit value=ȷ�� id=submit1 name=submit12>
      </form>
      <p><a href="stock.asp"><font style="font-size:9pt;" color="#0000FF">���ع�Ʊ���״���</font></a></p>
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
Jname=session("aqjh_name")
'sqlbbs="select * from [user] where username='" & Jname & "' "
sqlbbs="select ���� as userWealth from [�û�] where ����='" & Jname & "' "
set rsbbs=connbbs.execute(sqlbbs)
yin=rsbbs("userWealth")
if jin>yin then 
mess="��Ĺ�Ʊ�����˴���˵����������ô��Ǯ���������"
else
'sqlbbs="update [user] set userWealth=userWealth-"& jin & " where username='" & Jname & "'"
sqlbbs="update [�û�] set ����=����-"& jin & " where username='" & Jname & "'"
connbbs.execute sqlbbs
%>

<%       
sql="update �ͻ� set �ʽ�=�ʽ�+" & jin & " where �ʺ�='" & Jname & "'"
conn.execute sql
mess="���Ѿ���"&jin&"�����Ӵ������Ĺ�Ʊ�ʻ�"
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
      <p align=center><a href="stock.asp"><font style="font-size:9pt;" color="#0000FF">����</font></a></p>
    </td>
  </tr>
</table>
</body>
<%end sub%>