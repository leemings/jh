<%@ LANGUAGE=VBScript codepage ="936" %>
<%
if session("aqjh_jhdj")<20 then
	Response.Write "<script Language=Javascript>alert('��ʾ����Ա��������ϵͳ��Ҫ20���ſ��ԣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"& session("aqjh_name") &"'",conn
if rs("��Ա�ȼ�")>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������Ѿ��ǻ�Ա��,�뵽�ں�����ò��ã�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "SELECT * FROM q where d='δ��' and  b<now()-70000 order by b DESC",conn,3,3
do while not rs.eof 
	conn.Execute "update �û� set ״̬='���',�¼�ԭ��='���:\n����������������Ա��"&rs("e")&"��\n��Ϊ�㲻������Ǽ������˺ţ�!' where ����='"&rs("a")&"'"
	conn.Execute "update q set d='���' where a='"&rs("a")&"'"
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title><%=Application("aqjh_chatroomname")%>�û�ע��</title>
<link rel="stylesheet" href="../css.css">
</head>
<body bgcolor="#8EB4D9">
<center>
<table border=1 bgcolor="#FFFFFF" align=center width=400 cellpadding="5" cellspacing="10" background="../images/b2.gif" height="227">
<tr bgcolor="#FFFFFF" align="center">
      <td height="314" bgcolor="#8EB4D9"> <font color="#000000">�����������������������Ա��
      �����Ա</font>���취���<a href="../chat/hy.asp" target="_blank"><b><font color="#0000FF">����</font></b></a>��<br>
        <br>
        <font color="#FF0000"> �����Ա�Ƕ����ǽ�����������֧�֣�</font><br>
<br>
<table width="325" height="229">
<tr>
            <td height="236"> 
              <form method=POST action='regok.asp' onsubmit='return(check());' name=reg>
                <p align="center">������������<font color=yellow><b><%=session("aqjh_name")%></b></font>
  ���������룺    
                  <input type=password name=pass size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">   
                  <font color="#FF0000">*</font><br>   
                  <br>   
                  ��Ա�ȼ���    
                  <select name=hygrade onChange="javascript:js();" size=1 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">   
                    <option value="1" selected>һ��</option>   
                  </select>   
                  �� ����ʱ�䣺    
                  <select name=hymonth onChange="javascript:js();"  size=1 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">   
                    <option value="1" selected>01</option>   
                  </select>   
                  ��<br>   
                  <br>   
                  �����    
                  <input type=text name=money readonly size=5 maxlength="5" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="0">   
                  Ԫ <br>   
                </p>   
<p align="center">   
<input type=submit value=���� name="submit" style="border: 1px solid; font-size: 9pt; border-color:   
#000000 solid">   
<input type=button value=�ر� onClick="window.close()" name="button" style="border: 1px solid; font-size: 9pt; border-color:   
#000000 solid">   
</p>   
</form>   
</td>   
</tr>   
</table>   
</table>   
   
</center>   
</body>   
</html>   
<SCRIPT language="JavaScript">    
function js()   
{   
var hygrade=document.reg.hygrade.options.value;   
var hymonth=document.reg.hymonth.options.value;   
var money=10;   
if (hygrade=="1"){var money=10;}   
if (hygrade=="2"){var money=20;}   
if (hygrade=="3"){var money=40;}   
   
document.reg.money.value=money*hymonth;   
}   
function check()   
{   
var pass=document.reg.pass.value;   
   
if(pass=="" || pass==null){window.alert("�����������˰ɣ�����������������ʲô��Ա������");return false;}   
}   
</SCRIPT>