<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ����,����,w3 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
mygj=rs("����")
myfy=rs("����")
%>
<html>
<head>
<title>�ҵİ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css">
</head>
<body background="back.gif">
<p align="center">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="224" align="left">
  <tr> 
    <td valign="top" rowspan="3" width="416"> 
      <div align="center"><img src="dao.gif" width="40" height="15">δװ������һ��<img src="dao1.gif" width="40" height="15"> 
        <font color="#CC0000" face="��Բ"><a href="javascript:this.location.reload()">ˢ��</a></font></div> 
      <div align="center"> 
        <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
          <tr> 
            <td width="23%"> <div align="center">��Ʒ��</div></td>
            <td width="9%"> <div align="center">����</div></td>
            <td width="10%"> <div align="center">����</div></td>
            <td width="12%"> <div align="center">����</div></td>
            <td width="13%"> <div align="center">����</div></td>
            <td width="11%"> <div align="center">�;ö�</div></td>
            <td width="11%"> <div align="center">����</div></td>
            <td width="11%"> <div align="center">����</div></td>
          </tr>
          <%
if not(isnull(rs("w3"))) then
data1=split(rs("w3"),";")
data2=UBound(data1)
for y=0 to data2-1
	data3=split(data1(y),"|")
	rs.close
rs.open "select * FROM b WHERE a='" & data3(0) &"'",conn,3,3
If not(Rs.Bof and Rs.Eof) then
%>
          <tr align="center"> 
            <td width="23%" height="11"><img src='../jhjs/images/<%=rs("i")%>' alt='<%=rs("c")%>'>[<%=rs("a")%>]</td>
            <td width="9%" height="11"><%=rs("b")%></td>
            <td width="10%" height="11"> <%=data3(1)%> </td>
            <td width="12%" height="11"><%=rs("g")%></td>
            <td width="13%" height="11"><%=rs("f")%></td>
            <td width="11%"><%=rs("j")%></td>
            <td width="11%" height="11"><%=rs("k")%></td>
            <td width="11%" height="11"> <a href="wupin2.asp?name=<%=data3(0)%>"><font color="#0000FF">װ��</font></a> 
            </td>
          </tr>
          <%
end if
erase data3
next
erase data1
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
        </table>
        ע�⣺��װ����װ������Ʒ,������ͼƬ������˵����</div></td>
  </tr>
</table>
</body>
</html>
