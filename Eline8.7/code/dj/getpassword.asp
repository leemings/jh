<!--#include file="conn.asp"-->
<title>�û�ȡ������</title>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
<table width="100%" height="100%" border="0" cellspacing="1" bordercolor="#9C86AC">
  <tr>
    <td>
<%

if request("action")="getpass" then
	call getpass()
elseif request("action")="answer" then
	call showquestion()
else 
	call showuserform()
end if
sub showuserform()
%>
      <form action="getpassword.asp" method="post">
        <input type=hidden name=action value=answer>
        <div align="center">
          <center>
<table cellspacing=0 border=0 width="400" height="68" cellpadding="0" style="border-collapse: collapse">
                <tr> 
                  <td align=center height="1"> 
                  <b>ȡ������</b>����һ�������������û�����</td>
          </tr>
                <tr> 
                  <td height="126">
                  <p align="center"><font color="#FF0000">��������ע��ʱ���û���:<br>
                  <INPUT name=username type=text size=30 value="���������������û���" onclick="Javascript:this.value=''"></font></td>
                  </tr>
    	        <tr> 
                  <td align=center height="1"> 
                    <input type=submit name="submit" value="�ύ��������һ��">
                  </td>
                </tr>
                </table>
          </center></div>
</form>
<%
end sub

sub showquestion()
username=trim(request("username"))

if username="" then
response.write "<script language=javascript>alert('�������û���������');history.back(1);</script>"
	founderr=true
else
set rs=server.createobject("adodb.recordset")
	sql="select username,quesion from User where username='"&username&"'"
rs.open sql,conn,1,3
	if rs.bof and rs.eof then
response.write "<script language=javascript>alert('���û�����û���ڱ�վע�������ȷ�������Ƿ���ȷ������');history.back(1);</script>"
	founderr=true
	elseif rs("quesion")="" or isnull(rs("quesion")) then
response.write "<script language=javascript>alert('����ʧ�ܣ���ע��ʱû����дȡ������ʱ����Ҫ����Ϣ������');history.back(1);</script>"
	founderr=true
	end if
end if

if founderr=true then
	Response.End
	exit sub
else
end if


%>
<form action="getpassword.asp" method="post"> 
<input type=hidden name=action value=getpass>
<input type=hidden name=username value=<%=username%>>
           <div align="center">
             <center>
           <table cellspacing=0 border=0 width="400" bordercolor="#8DA2B7" height="137" cellpadding="0" style="border-collapse: collapse">
                <tr> 
                  <td colspan=2 align=center height="1"> 
                  <b>ȡ������</b>���ڶ������ش����⣩</td>
          </tr>
                <tr> 
                  <td  valign=middle width="121" height="67"> <b>������ʾ����:</b></td>
                  <td width="179" height="67"><font color="ff0000"><%=rs("quesion")%></font>��</td>
                </tr>
	            <tr> 
                  <td width="121" height="1"><b>ȡ�������:</b></td>
                  <td width="179" height="1"> 
                    <INPUT name=answer type=text size=30></td></tr>
	            <tr> 
                  <td colspan=2 align=center height="87"> 
                    <input type=submit name="submit" value="�ύ�����Ѿ���д��ȷ">
                  </td>
                </tr></table>
             </center></div>
</form>
<%
end sub

sub getpass()
answer=trim(request("answer"))
username=request("username")

if answer="" then
response.write "<script language=javascript>alert('����������Ĵ𰸣�����');history.back(1);</script>"
	founderr=true
else
	sql="select answer,PassWord from user where username='"&username&"'"
	set rs=conn.execute(sql)
	if rs("answer")<>answer then
response.write "<script language=javascript>alert('����ʧ�ܣ��������ش�ò���ȷ������');history.back(1);</script>"
		founderr=true
	end if
end if

if founderr=true then
	Response.End
	exit sub
end if


%>
<form action=Userlogin.asp method=post>
                
              <div align="center">
                <center>
                
              <table cellspacing=0 border=0 width="400" bordercolor="#8DA2B7" height="141" cellpadding="0" style="border-collapse: collapse">
                <tr align="center"> 
                  <td height="1" width="400" >
                  <b>ȡ������</b>���ɹ�ȡ�أ�</td>
    </tr>
                <tr> 
                  <td height="12" width="400">
                  <p align="center">&nbsp;&nbsp;��ϲ�������ɹ�ȡ�����룡</td>
          </tr>
	            <tr> 
                  <td height="58" width="400" align="right">
                  <p align="center"><b>�û���:<font color="#ff0000"><%=username%></font></b></p>
                  <p align="center"></p>
                  <p align="center">
                  <b>��&nbsp;&nbsp;��:<font color="#ff0000"><%=rs("PassWord")%></font></b></td>
                </tr>
	            <tr> 
                  <td height="35" width="400">
                  <p align="center"><b>˵&nbsp;&nbsp;��:</b>�����Ʊ����������룡</td>
                </tr>
                <tr align="center"> 
                  <td height="36" width="400"> 
                    <input type=button name=login value=�ɹ�ȡ�����룬�رմ˴��� onclick=window.close()>
                  </td>
    </tr>  
    </table></center></div>
       </form>
<%
end sub
%></td>
  </tr>
</table>
  </center></div>