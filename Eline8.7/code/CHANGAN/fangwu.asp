<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
my=sjjh_name
%>
<!--#include file="data1.asp"--><%
sql="SELECT * FROM ���� where  ����='һ�㷿��'"
Set Rs=conn.Execute(sql)
%>
<html>

<head>
<LINK 
href="../pic/css.css" rel=stylesheet>
</head>

<body bgcolor="#FFFDDF">
<table border=1 bgcolor="#588878" align=center width=98% cellpadding="0" cellspacing="1" bordercolor="#65251C">
  <tr bgcolor="#006633"> 
    <td height="15" colspan="6" bgcolor="#588878"> 
      <p align="center"> <img src="aqwjy_pic.jpg" width="311" height="93" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <img src="aqwjy_zi.gif" width="126" height="93"> 
    
    
</table>
<table border=1 bgcolor="#FFFFCC" align=center width=98% cellpadding="0" cellspacing="1" bordercolor="#65251C">
  <tr> 
    <td height="30" colspan="7" bgcolor="#FFFFFF"> 
      <p align=center class="p9"><font style="FONT-SIZE: 9pt" color="#FF3333"><marquee><font color="#990099"> 
        <a href="../welcome.asp">���ؽ���</a> ��ӭ</font><%=name%><font color="#990099">���ٷ������ģ�˭�������Լ��ļң���������������ĸ���Ҫ��(������������ֻ�ܵõ�ԭ����1/5�ļ۸�)</font></marquee></font></p>  
  <tr>   
    <td height="21" colspan="7" bgcolor="#FFFFFF">   
      <table width="100%" border="1" cellspacing="1" cellpadding="1" bordercolor="#6633FF">  
        <tr>   
          <td bgcolor="#65251C">   
            <div align="center"><a href=fangwu.asp><font color="#FFFFFF">һ�㷿��</font></a></div>  
          </td>  
          <td bgcolor="#65251C">   
            <div align="center"><a href=fangwu3.asp><font color="#FFFFFF">�߼���Ԣ</font></a></div>  
          </td>  
          <td bgcolor="#65251C">   
            <div align="center"><a href=fangwu4.asp><font color="#FFFFFF">��԰��</font></a></div>  
          </td>  
          <td bgcolor="#65251C">   
            <div align="center"><a href=fangwu5.asp><font color="#FFFFFF">��������</font></a></div>  
          </td>  
        </tr>  
      </table>  
  <tr bordercolor="#000000">   
    <td height="15" width="14%" bgcolor="#65251C" bordercolor="#65251C">   
      <div align="center"><font size="2" color="#FFFFFF">��������</font></div>  
    </td>  
    <td height="15" width="9%" bgcolor="#65251C" bordercolor="#65251C">   
      <div align="center"><font size="2" color="#FFFFFF">�ۼ�</font></div>  
    </td>  
    <td height="15" width="11%" bgcolor="#65251C" bordercolor="#65251C">   
      <div align="center"><font size="2" color="#FFFFFF">����</font></div>  
    </td>  
    <td height="15" colspan="2" bgcolor="#65251C" bordercolor="#65251C">   
      <div align="center"><font size="2" color="#FFFFFF">���ƺ���</font></div>  
    </td>  
    <td height="15" width="19%" bgcolor="#65251C" bordercolor="#65251C">   
      <div align="center"><font size="2" color="#FFFFFF">����/����</font></div>  
    </td>  
    <td height="15" width="25%" bgcolor="#65251C" bordercolor="#65251C">   
      <div align="center"><font size="2" color="#FFFFFF">����</font></div>  
    </td>  
  </tr>  
  <%   
do while not rs.eof and not rs.bof   
%>   
  <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';">   
    <td width="14%" class="calen-curr" height="15"><font size="2">   
      <center>  
        <%=rs("����")%>   
      </center>  
      </font>   
    <td width="9%" class="calen-curr" height="15"><%=rs("�ۼ�")%>   
      <div align="center"></div>  
    <td width="11%" class="ct-def4" height="15">   
      <div align="center"><font size="2"><%=rs("λ��")%></font></div>  
    </td>  
    <td colspan="2" class="ct-def4" height="15"><%=rs("�ֵ�")%> <font color=red><%=rs("id")%></font><font size="2">��</font></td>  
    <td width="19%" class="calen-curr" height="15">   
      <div align="center"><font size="2"><%=rs("����")%> / </font><font color="ff00ff"><%=rs("����")%></font></div>  
    </td>  
    <td width="25%" class="calen-curr" height="15"><font size="2">   
      <center>  
        <%if rs("����")<>name and rs("����")="��" then%>  
        <a href=fangwu1.asp?id=<%=rs("id")%>><font color="0000ff">����</font></a>  
        <%end if%>  
        <%if rs("����")=sjjh_name then%>  
        <a href=fangwu2.asp?id=<%=rs("id")%>><font color=red>����</font></a>  
        <%end if%>  
        <%if rs("����")<>"��" and rs("����")<>name then%>  
        �÷�������  
        <%end if%>  
      </center>  
      </font></td>  
  </tr>  
  <%   
rs.movenext   
loop   
  
	set rs=nothing	  
	conn.close  
	set conn=nothing  
%>   
</table>  
</html>