<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
lb=LCase(trim(Request.QueryString("lb")))
if lb="" then lb="ҩƷ"
%>
<html>

<head>
<title>���ֽ�����Ʒ��������</title>
<style></style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center">��Ʒ��������<br>
  <br>
  <a href="wpadmin.asp?lb=����">����</a> <a href="wpadmin.asp?lb=����">����</a> 
  <a href="wpadmin.asp?lb=װ��">װ��</a> <a href="wpadmin.asp?lb=����">����</a> 
  <a href="wpadmin.asp?lb=˫��">˫��</a> <a href="wpadmin.asp?lb=ͷ��">ͷ��</a> 
  <a href="wpadmin.asp?lb=����">����</a> <a href="wpadmin.asp?lb=����">����</a> 
  <a href="wpadmin.asp?lb=ҩƷ">ҩƷ</a> <a href="wpadmin.asp?lb=��ҩ">��ҩ</a> 
  <a href="wpadmin.asp?lb=��Ƭ">��Ƭ</a> <a href="wpadmin.asp?lb=�ʻ�">�ʻ�</a> 
  <a href="wpadmin.asp?lb=����">����</a> <a href="wpadmin.asp?lb=��԰">��԰</a> 
<a href="wpadmin.asp?lb=��ͨ">��ͨ</a> <a href="wpadmin.asp?lb=ħ��">ħ��</a> 
<a href="wpadmin.asp?lb=����">����</a> 
</p>
<table width="600" border="1" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000"
bgcolor="#006699">
  <tr> 
    <td colspan="4" height="11" bgcolor="#336633"> 
      <p align="center"><font color="#FFFF00">�� �� <b>[<%=lb%>]</b> �� ��</font></p>
    </td>
    <td height="11" bgcolor="#336633" width="84"> 
      <div align="center"><a href="addwp.asp?lb=<%=lb%>"><font color=ffffff>������Ʒ</a></div>
    </td>
  </tr>
  <tr> 
    <td width="82" height="20" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">��Ʒ��</font></div>
    </td>
    <td width="63" height="20" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">�� ��</font></div>
    </td>
    <td width="259" height="20" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">�� ��</font> </div>
    </td>
    <td width="94" height="20" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">�� ��</font> </div>
    </td>
    <td width="84" height="20" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">�� ��</font> </div>
    </td>
  </tr>
  <%
sql="SELECT * FROM b where b='"&lb&"' order by b,h "
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
  <tr bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#85C2E0';"> 
    <td width="82" height="23"> 
      <div align="center"><font color="#FFFFFF" size="2"><%=rs("a")%> </font></div>
    </td>
    <td width="63" height="23"> 
      <div align="center"><font color="#FFFFFF" size="2"><%=rs("b")%></font></div>
    </td>
    <td width="430" height="23"> 
      <div align="left"><font color="#FFFFFF" size="2">
      <%select case lb
      	case "��Ƭ","����","�ʻ�"
	      	Response.Write "˵��:"&rs("c")
			if lb="����" then
				Response.Write " ����:"&rs("k")
			end if
	      	if lb="����" or lb="�ʻ�" then
   		      	Response.Write " ͼ:<img src='../hcjs/jhjs/images/"&rs("i")&"'>"
	      	end if
		case "��ҩ","����","ҩƷ"
			Response.Write "������"&abs(rs("d"))&" ������"&abs(rs("e"))
		case else
			Response.Write  "������"&abs(rs("f"))&" ����:"&abs(rs("g"))
			if lb="����" or lb="����" then
			Response.Write " ����:"&rs("k")
			end if
	      	if lb="����" or lb="����" or lb="����" or lb="˫��" or lb="װ��" or lb="ͷ��" or lb="����" then
   		      	Response.Write " �;�:"&rs("j")&" ͼ:<img src='../hcjs/jhjs/images/"&rs("i")&"'>"
	      	end if
      	end select
      %>
        
        </font></div>
    </td>
    <td width="94" height="23"> 
      <div align="right"><font color="#FFFFFF" size="2"><%=rs("h")%><%if lb<>"��Ƭ" then%> ��<%else%> Ԫ<%end if%></font></div>
    </td>
    <td width="84" height="23"> 
      <div align="center"><a
href="modiwp.asp?wupinid=<%=rs("id")%>"><font color="#FFFFFF" size="-1">�޸�</font></a> 
        <font size="-1"><a href="del.asp?id=<%=rs("id")%>"><font color="#FFFFFF">ɾ��</font></a></font></div>
    </td>
  </tr>
  <%
rs.movenext
loop
conn.close
%>
</table>
</body>
</html>