<%
Response.Expires=0
dim conn,rs,userconn,users
username=Session("aqjh_name")
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
CurPage = 1
Else
CurPage = CINT(Request.QueryString("CurPage"))
End If
if Session("aqjh_name")=""  then
  Response.Redirect "../error.asp?id=440"
else
%>
<html>
<head>
<title>��ԩ״ֽһ��</title>
<LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css><!--td {  font-family: ����; font-size: 9pt}body {  font-family: ����; font-size: 9pt}select {  font-family: ����; font-size: 9pt}A {text-decoration: none; font-family: "����"; font-size: 9pt}A:hover {text-decoration: underline; color: #CC0000; font-family: "����"; font-size: 9pt} .big {  font-family: ����; font-size: 12pt}
--></style>

</head>
<body oncontextmenu=self.event.returnValue=false leftmargin="0" topmargin="0" bgcolor="#3a4b91">
<table width="650" border="0" cellspacing="2" cellpadding="2" align="center" bgcolor="#0066CC">
  <tr> 
    <td colspan="2" height="26">&nbsp;<a href=index.asp onClick='javascript:history.back()'>�� 
      ��</a></td>
  </tr>
</table>   
<form action=over.asp method=get>
<%
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
rs.open "SELECT * FROM sy ORDER BY id DESC",Conn,1,1
if not rs.eof and not rs.bof then
RS.PageSize=15
Dim TotalPages
TotalPages = RS.PageCount

If CurPage>RS.Pagecount Then
CurPage=RS.Pagecount
end if

RS.AbsolutePage=CurPage

rs.CacheSize = RS.PageSize

Dim Totalcount
Totalcount =INT(RS.recordcount)
%>
  <table width="650" border="1" cellspacing="0" cellpadding="2" align="center" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr> 
      <td><font size="2" color="#000000">�� ����<%=Totalcount%>��״ֽ����<%=TotalPages%>ҳ��ĿǰΪ��<%=CurPage%>ҳ</font></td>
      <td align="right"><font size="2" color="#000000"> <a href="index.asp"><font color="#FF0000">����״ֽ</font></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=manage.asp?owner=<%=owner%>>[ˢ��]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=<%=Curpage-1%>>[��һҳ]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=<%=Curpage+1%>>[��һҳ]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=1>[��ҳ]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=<%=TotalPages%>>[βҳ]</a>&nbsp;</font></td>
    </tr>
  </table>             
  <table width=650 border=0 cellspacing=1 cellpadding=0 align="center" >
    <tr height=22 bgcolor="#0066CC"> 
      <td height="25" width="334"><font size="2"><b><font color="#000000">&nbsp;&nbsp;&nbsp;</font></b><font color="#FFFFFF">״ֽ����</font></font></td>
      <td align="center" colspan="2" height="25"><font size="2" color="#FFFFFF">�ܺ���</font></td>
      <td align="center" height="25" width="100"><font size="2" color="#FFFFFF">������ 
        </font></td>
      <td align="center" height="25" width="93"><font size="2" color="#FFFFFF">���н��</font></td>
    </tr>
    <%I=0             
p=RS.PageSize*(Curpage-1)             
do while (Not RS.Eof) and (II<RS.PageSize)             
p=p+1%> 
    <tr bgcolor="#CCCCCC"> 
      <td height="25" width="334"><font size="2"><a href=rootdisp.asp?ID=<%=rs("id")%>><%=rs("����")%></a> 
        </font></td>
      <td colspan="2" height="25"><b><font size="2"><%=left(rs("ԭ��"),10)%></font></b></td>
      <td width="100" height="25"><b><font size="2"><%=left(rs("����"),10)%></font></b></td>
      <td width="93" height="25"><b><font size="2"><%if rs("���")="N" then%>δ����<% end if %><%if  rs("���")="0" then %>������<% end if %><% if  rs("���")="1" then %>����<% end if %></font></b></td>
    </tr>
    <%rs.movenext             
II=II+1             
loop%> 
  </table>
  <%StartPageNum=1             
do while StartPageNum+15<=CurPage             
StartPageNum=StartPageNum+15             
Loop             
                 
EndPageNum=StartPageNum+14             
             
If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount             
%> 
  <table width="650" border="1" cellspacing="0" cellpadding="3" align="center" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr> 
      <td> 
        <div align="center">�� ҳ�Σ� <font color="#CC0000"><%=CurPage%></font>/<%=TotalPages%> 
          ��ÿҳ��<font color="#CC0000"><%=RS.PageSize%></font>�� </div>
      </td>
      <td align="right"> 
        <div align="center">ҳ���� <a href="manage.asp?owner=<%=owner%>&CurPage=<%=StartPageNum-1%>"><<</a> 
          <% For I=StartPageNum to EndPageNum               
      if I<>CurPage then %> <a href="manage.asp?owner=<%=owner%>&CurPage=<%=I%>">[<%=I%>]</a> 
          <% else %> <font color="#CC0000"><b><%=I%></b></font> <% end if %> <% Next %> 
          <% if EndPageNum<RS.Pagecount then %> <a href="manage.asp?owner=<%=owner%>&CurPage=<%=EndPageNum+1%>">[����...]</a> 
          <%end if%> </div>
    </tr>
  </table>             
  <%else%> <br>
  <div align="center"><font color="#FFFFFF">��ʱ��״ֽ����</font><font color="#FF0001"><a href="index.asp">������״ֽ</a></font> 
    <%   
end if%> </div>
</form>             
   
<% end if%> 
<div align="center">
  <hr width="80%">
</div>
</body>             
</html>    
<%   
rs.close   
set rs=nothing   
%>