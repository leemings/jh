<!--#include file="conn.asp"-->
<!--#include file="chkmaster.asp"-->

<% 
if session("yx8_mhjh_username")="" then
response.redirect "../login.asp"
else
if session("yx8_mhjh_username")<>Application("yx8_mhjh_admin") then
call endinfo("��û���ʸ����")
else %>
<html>
<head>
<title>�ͻ�����</title>
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
          <td nowrap align=center><font style="font-size:9pt;" color=#ffffff><a href="stock.asp"><font color=white>��Ʊ��ҳ</font></a>  ----------------------<a href="userma.asp"><font color=white>�ͻ�����</font></a>----------------------</font><a href="gpma.asp"><font color=white>��Ʊ����</font></a></td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  <tr  bgcolor="#0099FF"> 
   <form name="form1" method="post" action=userma.asp>
    <td colspan="4" align=center height=18> 
        <input type="text"  name=userma><input type="submit" name="Submit" value="��ѯ�ͻ�">
      </td>
       </form>
  </tr>
  <tr align="center" > 
    
    <td > 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;" color=#ffffff>�ͻ�</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td > 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;" color=#ffffff>�ʽ�</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td > 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;" color=#ffffff>״̬</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
	<td > 
      <table  cellspacing=0 cellpadding=1 border=0>
        <tbody> 
        <tr> 
          <td><font style="font-size:9pt;" color=#ffffff>����</font></td>
        </tr>
        </tbody> 
      </table>
    </td>
     </tr>
	 
  <%      if request("userma")=""  then     
sql= "select * from �ͻ� order by �ʽ� desc"         
set rs=conn.execute(sql)         
DO UNTIL RS.EOF  %>
  <tr> 
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><%=rs("�ʺ�")%></font></td>
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p"><%=rs("�ʽ�")%></span></font></td>
    <td align=center bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p"><%if rs("����")=1 then%>����<%else%>δ��<%end if%> </span></font></td>
	<td bgcolor="#FFFFFF"> <div align="center"><font style="font-size:9pt;" color="#000066"><a href="userma.asp?userma=<%=rs("�ʺ�")%>">�޸�</a></font></div>
    </td>
  </tr>
  <%         
  rs.MoveNext         
  Loop         
else
  if request("action")="save" then
  sql= "select * from �ͻ� where �ʺ�='"&request("userma")&"'"         
 set rs=server.createobject("adodb.recordset")
 rs.open sql,conn,1,3
rs("�ʽ�")=request.form("money")
rs("����")=trim(request.form("lock"))
rs.update
rs.close
call endinfo("�û���Ϣ�Ѿ��޸����!")
else
sql= "select * from �ͻ� where �ʺ� like '%"&request("userma")&"%'"         
set rs=conn.execute(sql)         
 if not rs.eof then
    do while not rs.eof
 %>
  <tr> 
  <form method=post  action=userma.asp?userma=<%=rs("�ʺ�")%>&action=save>
    <td bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><%=rs("�ʺ�")%></font></td>
    <td align=center bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p"><input type=text name=money value="<%=rs("�ʽ�")%>"></span></font></td>
    <td align=center bgcolor="#FFFFFF" ><font style="font-size:9pt;" color="#000066"><span class="p"><select name="lock" size=1 >
<option value="<%=rs("����")%>" selected><%if rs("����")=1 then%>����<%else%>δ��<%end if%></option>
 
<option value="<%=1-rs("����")%>"><%if rs("����")=1 then%>����<%else%>����<%end if%></option>

 	</select> </span></font></td>
	<td bgcolor="#FFFFFF"> <div align="center"><input type=submit name=submit value=ִ���޸�></font></div>
    </td>
	</form>
  </tr>
  <%rs.movenext
  loop
  rs.close
  else
  call endinfo("�������ҵ��û�������!")
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
CloseDatabase		'�ر����ݿ�

sub endinfo(message) 
'-------------------------------��Ϣ��ʾ-------------------------------
%>
<table align=center cellspacing=1 cellpadding="3" width="97%">
	<tr>
		<td align="center"><b>��Ϣ��ʾ</b></td>
	</tr>
	<tr>
		<td align=center height=70 bgcolor="#CCCCCC"><%=message%><br></td>
	</tr>
	<tr>
		<td align=center height=26 bgcolor="#0099FF"><a href="userma.asp">����</a></td>
	</tr>
</table>
<%end sub
%>
