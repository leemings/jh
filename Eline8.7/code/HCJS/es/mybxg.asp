<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
if Session("sjjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();}</script>"
	Response.End 
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
rs.open "delete * from s where m<now()-5 and c<>'���չ�'",conn,3,3
rs.open "Select * From s where c='���չ�' and b="& sjjh_id &" Order by m DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>

<head>
<title><%=Application("sjjh_chatroomname")%>-�ҵı��չ�</title>
<link rel="stylesheet" href="style.css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body leftmargin="0" topmargin="0" bgcolor="#66CCCC">
<table border="0" height="24" width="91%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="15" width="100%" bgcolor="#669999"><font color="#669966"> <font color="#FFFFFF"><b>���ͱ��չ�(</b><font color="#000000">���������Ʒ�ǲ��ᶪʧ�ģ�����Ҫ��5���շѣ�</font><b>��</b></font></font></td>
</tr>
</tbody>
</table>
<div align="center">
<table width="91%" align="center" cellspacing="0" border="0"
cellpadding="0">
<tr>
<td width="100%">
<table border="1" cellspacing="1" cellpadding="0" width="720"
bordercolorlight="#EFEFEF" height="8">
<tr bgcolor="#FFFFFF">
            <td width="9%" height="16"> 
              <div align="center"><font color="#FF6600"> ��Ʒ��</font></div>
</td>
<td width="7%" height="16">
<div align="center"><font color="#FF6600"> ӵ����</font></div>
</td>
            <td width="6%" height="16"> 
              <div align="center"><font color="#FF6600"> ����</font></div>
</td>
            <td width="21%" height="16"> 
              <div align="center"><font color="#FF6600"> ʱ �� </font></div>
</td>
            <td width="36%" height="16"> 
              <div align="center"><font color="#FF6600">˵ ��</font></div>
</td>
            <td width="9%" height="16"> 
              <div align="center"><font color="#FF6600"> ȡ������</font></div>
</td>
            <td width="12%" height="16"> 
              <div align="center"><font color="#FF6600">����</font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
<tr bgcolor="#3399CC" onmouseout="this.bgColor='#3399CC';"onmouseover="this.bgColor='#DFEFFF';">
            <td width="9%" height="26"> 
              <div align="center"> <font color="#0000FF"><%=rs("a")%></font>
</div>
</td>
<td width="7%" height="26">
<div align="center"> <%=rs("b")%> </div>
</td>
            <td width="6%" height="26"> 
              <div align="center"> <%=rs("j")%> </div>
</td>
            <td width="21%" height="26"> 
              <div align="center"> <%=rs("m")%> </div>
</td>
            <td width="36%" height="26"> 
              <div align="left"></div>
<%=rs("k")%></td>
<form method=POST action='mybxg1.asp?id=<%=rs("id")%>'>
            <td width="9%" height="26"> 
              <div align="center"> <font color="#FFFFFF">
                <select name="wpsl">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
                </font></div>
</td>
            <td width="12%" height="26"> 
              <div align="center"><input type="SUBMIT" name="Submit"  value="����"></div>
</td>
</form>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
pa=rs.pagecount
mu=musers()
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<table border="0" cellspacing="1" cellpadding="1" width="100%" bordercolorlight="#EFEFEF">
<tr>
<td align="left" width="37%" height="2">[��<font color="red"><b><%=pa%></b></font>ҳ][<font
color="red"><b><%=mu%></b></font>����Ʒ]</td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="mybxg.asp?page=<%=page-1%>">��һҳ</a>][��<%=page%>ҳ][<a href="mybxg.asp?page=<%=page+1%>">��һҳ</a>]</div>
</td>
</tr>
</table>
</table>
</div>

<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As ���� from s where c='���չ�' and b="& sjjh_id)
musers=tmprs("����")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>
</body>
