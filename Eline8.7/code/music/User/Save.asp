<%PageName="UserSave"%>
<!--#include file="conn.asp"-->
<!--#include file="function.asp"-->
<!--#include file="Star.INC"-->
<%
founderr=false
if isnull(session("password")) or session("password")="" or isnull(session("username")) or session("username")="" then
	errmsg=errmsg+"<br>"+"����δ��½��"
	founderr=true
end if
if trim(request.form("password"))="" or Len(request.form("password"))>10 then
	errmsg=errmsg+""+"��������������(���Ȳ��ܴ���10)��"
	founderr=true
else
	password=trim(request.form("password"))
end if
if password<>request.form("password2") then
	errmsg=errmsg+""+"������������ȷ�����벻һ�¡�"
	founderr=true
end if
if IsValidEmail(trim(request.form("Email")))=false then
	errmsg=errmsg+""+"����Email�д���"
	founderr=true
else
	Email=trim(request.form("Email"))
end if



if founderr=true then
	call error()
else
	set rs=server.createobject("adodb.recordset")
	sql="select * from [user] where username='"&session("username")&"' and password='"&session("password")&"'"
	rs.open sql,conn,1,3
	if rs.eof then
		errmsg="<br>"+"�Բ����㲻�Ǳ�վ�û���������ע�ᡣ"
		founderr=true
	else
		rs("username")=session("username")
		rs("password")=password
		rs("email")=email

		if trim(request.form("Sex"))="" then
			rs("Sex")=null
		else
			rs("Sex")=trim(request.form("Sex"))
		end if

		if trim(request.form("Tel"))="" then
			rs("Tel")=null
		else
			rs("Tel")=trim(request.form("Tel"))
		end if

		if trim(request.form("Name"))="" then
			rs("Name")=null
		else
			rs("Name")=trim(request.form("Name"))
		end if

		if trim(request.form("Address"))="" then
			rs("Address")=null
		else
			rs("Address")=trim(request.form("Address"))
		end if

		if trim(request.form("Youbian"))="" then
			rs("Youbian")=null
		else
			rs("Youbian")=trim(request.form("Youbian"))
		end if

		if trim(request.form("Shenfenzheng"))="" then
			rs("Shenfenzheng")=null
		else
			rs("Shenfenzheng")=trim(request.form("Shenfenzheng"))
		end if

		rs.update
		session("password")=password
	end if
	rs.close

	if founderr=true then
		call error()
	else
%>
<!--#include file="Top.asp"-->
<div align="center">         
<table border="0" cellpadding="0" cellspacing="0" width="500"> 
<tr> 
<td width="100%" height="12" colspan="2"></td> 
</tr>           
<tr> 
<td width="70%"></td> 
<td width="30%" height="17" background="../images/wanbj.gif" align="right">�� �� �� ��&nbsp;&nbsp;&nbsp; </td> 
</tr> 
<tr> 
<td width="100%" colspan="2"><img border="0" src="../images/xiaod.gif" width="1" height="1"></td> 
</tr> 
<tr> 
<td width="100%" colspan="2" bgcolor="#56B0F4"><img border="0" src="../images/xiaod.gif" width="1" height="2"></td> 
</tr>      
</table>               
</div>        
<div align="center">        
  <center>        
  <table border="0" cellspacing="0" width="502">        
    <tr>        
      <td width="100%">        
        <div align="center">       
          <table border="0" cellpadding="4" cellspacing="1" width="502"> 
           <form method=post action="ChkLogin.asp">       
            <tr>       
              <td width="180" bgcolor="#9FD6F8" align="right">��¼ID��</td>       
              <td width="320" bgcolor="#9FD6F8">&nbsp;<%=session("username")%></td>       
            </tr>   
            <tr>      
              <td width="180" bgcolor="#b4def8" align="right">E-mail��</td>     
              <td width="320" bgcolor="#b4def8">&nbsp;<%=Email%></td>     
            </tr>     
            <tr>     
              <td width="180" bgcolor="#9FD6F8" align="right">�ƺ���</td>     
              <td width="320" bgcolor="#9FD6F8">&nbsp;<%=request("Sex")%></td>    
            </tr>    
            <tr>    
              <td width="180" bgcolor="#b4def8" align="right">��ʵ������</td>    
              <td width="320" bgcolor="#b4def8">&nbsp;<%=request("Name")%></td>    
            </tr>    
            <tr>    
              <td width="180" bgcolor="#9FD6F8" align="right">���֤�ţ�</td>    
              <td width="320" bgcolor="#9FD6F8">&nbsp;<%=request("Shenfenzheng")%></td>    
            </tr>    
            <tr>    
              <td width="180" bgcolor="#b4def8" align="right">��ϵ�绰��</td>    
              <td width="320" bgcolor="#b4def8">&nbsp;<%=request("Tel")%></td>    
            </tr>
            <tr>    
              <td width="180" bgcolor="#9FD6F8" align="right">��ϸ��ַ��</td>    
              <td width="320" bgcolor="#9FD6F8">&nbsp;<%=request("Address")%></td>    
            </tr>    
            <tr>    
              <td width="180" bgcolor="#b4def8" align="right">�������룺</td>    
              <td width="320" bgcolor="#b4def8">&nbsp;<%=request("Youbian")%></td>    
            </tr>
            <tr>    
              <td width="500" bgcolor="#67BDF1" align="center" colspan="2"><INPUT type=button name=back value="�� �� �� ҳ" onclick="javascript:window.open('../index.asp','_self','')"></td>     
            </tr>   
          </table>     
        </div>     
      </td>     
    </tr>     
    <tr>     
      <td width="100%" height="12"></td>     
    </tr>     
  </table>     
  </center>     
</div>             
<!--#include file="Bottom.asp"-->
<%
	end if
	set rs=nothing
end if
conn.close
set conn=nothing
%>

