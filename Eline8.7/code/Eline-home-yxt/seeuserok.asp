<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
tiaojian=Request.Form("tiaojian")
show=trim(Request.Form("show"))
%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>�û����ϲ鿴����</title>
<style type="text/css">
<!--
p            { line-height: 20px; font-size: 9pt }
table        { font-size: 9pt }
a:link       { color: #FF9900; text-decoration: none }
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" vlink="#FF9900" topmargin="0"
leftmargin="0" background="../jhimg/bk_hc3w.gif">
<p align="center"> <font color="#CC0000" face="��Բ"><a href="javascript:this.location.reload()">ˢ��</a></font> 
  <br>
  ��һЩ�������������ˣ������������޸ģ�<br>
  <font color="#FF0000"><b><%=tiaojian%></b></font> <br>
 <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
	if show<>"" then
		rs.open "SELECT * FROM �û� where "& tiaojian &" order by "&show,conn
	else
		rs.open "SELECT * FROM �û� where "& tiaojian &" order by lasttime",conn
	end if
%>
<table border="0" width="500" cellspacing="0" cellpadding="0"
background="../jhmp/bg.gif" align="center">
  <tr align="center">
    <td background="../jhmp/top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1"><%=Application("sjjh_chatroomname")%></font></b></font></td>
</tr>
<tr align="center">
<td>
      <table width="485" border="1" cellpadding="0" cellspacing="0" height="13">
        <tr> 
          <td width="28" height="10"> 
            <div align="center"><font color="#FFFFFF">ID</font></div>
          </td>
          <td width="47" height="10"> 
            <div align="center"><font color="#FFFFFF">����</font></div>
          </td>
          <td width="25" height="10"> 
            <div align="center"><font color="#FFFFFF">�Ա�</font></div>
          </td>
          <td width="63" height="10"> 
            <div align="center"><font color="#FFFFFF">����</font></div>
          </td>
          <td width="54" height="10"> 
            <div align="center"><font color="#FFFFFF">���</font></div>
          </td>
          <td width="82" height="10"> 
            <div align="center"><font color="#FFFFFF">����½ʱ��</font></div>
          </td>
          <td width="58" height="10"> 
            <div align="center"><font color="#FFFFFF">�����ȼ�</font></div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="10"> 
            <div align="center"><font color="#FFFF00"><b><%=show%></b></font></div>
          </td>
          <%end if%>
          <td width="35" height="10"> 
            <div align="center"><font color="#FFFFFF">��¼</font></div>
          </td>
        </tr>
        <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
        <tr> 
          <td width="28" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
          </td>
          <td width="47" height="30"> 
            <div align="center"><a href="SHOWUSER.asp?username=<%=rs("����")%>"><font color="#FF9900"><%=rs("����")%></font></a></div>
          </td>
          <td width="25" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("�Ա�")%></font></div>
          </td>
          <td width="63" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("����")%></font></div>
          </td>
          <td width="54" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("���")%></font></div>
          </td>
          <td width="82" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("lasttime")%></font></div>
          </td>
          <td width="58" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("�ȼ�")%></font></div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="30"> 
            <div align="center"><font color="#FFFF00"><b><%=rs(show)%></b></font></div>
          </td>
          <%end if%>
          <td width="35" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("times")%></font></div>
          </td>
        </tr>
        <%
rs.movenext
loop
conn.close
%>
      </table>
</td>
</tr>
<tr align="center">
    <td background="../jhmp/top3.gif" width="500" height="2"><font color="#FFFFFF"><%=Application("sjjh_chatroomname")%>�����޸�</font></td>
</tr>
</table>
<div align="center"><font color="#000000">��������:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">��</font> 
</div>
</body>
</html>