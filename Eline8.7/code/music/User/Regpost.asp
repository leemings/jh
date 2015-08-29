<%PageName="UserRegPost"%>
<!--#include file="conn.asp"-->
<!--#include file="function.asp"-->
<!--#include file="Star.INC"-->
<%
founderr=false

if request.form("UserName")="" or len(request.form("UserName"))>20 then
	errmsg=errmsg+""+"用户名输入错误(未输入或长度超过了20个字节)。<br>"
	founderr=true
else
	UserName=trim(request.form("UserName"))
end if
if request.form("password")="" or Len(request.form("password"))>20 then
	errmsg=errmsg+""+"请输入您的密码(长度不能大于20)。<br>"
	founderr=true
else
	password=request.form("password")
end if

if password<>request("password2") then
	errmsg=errmsg+""+"您输入的密码和确认密码不一致。"
	founderr=true
end if
if IsValidEmail(trim(request.form("Email")))=false then
	errmsg=errmsg+""+"您的Email有错误。"
	founderr=true
else
	Email=trim(request.form("Email"))
end if



if founderr=true then
	call error()
else
	set rs=server.createobject("adodb.recordset")
	sql="select * from [user] where username='"&username&"'"
	rs.open sql,conn,1,3
	if not rs.eof or username=WebName then
		errmsg=""+"对不起，您输入的用户名已经被注册，请重新输入。"
		founderr=true
	else
		rs.addnew
		rs("username")=username
		rs("password")=password
		rs("email")=email
		if request.form("Address")<>"" then rs("Address")=request.form("Address")
		if request.form("Tel")<>"" then rs("Tel")=request.form("Tel")
		if request.form("Sex")<>"" then rs("Sex")=request.form("Sex")
		if request.form("Youbian")<>"" then rs("Youbian")=request.form("Youbian")
		if request.form("Name")<>"" and request.form("Shenfenzheng")<>"" then
			rs("Name")=request("Name")
			rs("Shenfenzheng")=request("Shenfenzheng")
		end if
		rs("addDate")=NOW()
		rs.update
	end if
	rs.close

	if founderr=true then
		call error()
	else
%>
<!--#include file="Top.asp"-->
<div align="center">
  <table width="766" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#FFFFFF"> <br>
        <table border="0" cellspacing="0" cellpadding="10" align="center" bgcolor="#9CD7FF" width="467">
          <tr> 
            <td bgcolor="#E6F4FF">
              <table border="0" cellspacing="0" width="502">
                <tr> 
                  <td width="100%"> 
                    <div align="center"> 
                      <table border="0" cellpadding="4" cellspacing="1" width="502">
                        <tr align="center"> 
                          <td colspan="2" bgcolor="#9FD6F8"><font color="#FF0000">注册成功</font></td>
                        </tr>
                        <form method=post action="ChkLogin.asp">
                          <tr> 
                            <td width="180" bgcolor="#9FD6F8" align="right">登录ID：</td>
                            <td width="320" bgcolor="#9FD6F8">&nbsp; 
                              <input type=text name=username size=26 class=flat1 style="background-color: #9FD6F8; color: #ff0000; border: 0 solid #000000" value="<%=username%>">
                              <input type=password name=password size=1 class=flat1 style="color: #9FD6F8; background-color: #9FD6F8; border: 0 solid #000000" value="<%=password%>">
                            </td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#b4def8" align="right">E-mail：</td>
                            <td width="320" bgcolor="#b4def8">&nbsp;<%=Email%></td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#9FD6F8" align="right">称呼：</td>
                            <td width="320" bgcolor="#9FD6F8">&nbsp;<%=request.form("Sex")%></td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#b4def8" align="right">真实姓名：</td>
                            <td width="320" bgcolor="#b4def8">&nbsp; 
                              <%if request.form("Name")="" then%>
                              未填写 
                              <%else%>
                              <%=request.form("Name")%> 
                              <%end if%>
                            </td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#9FD6F8" align="right">OICQ：</td>
                            <td width="320" bgcolor="#9FD6F8">&nbsp; 
                              <%if request.form("Shenfenzheng")="" then%>
                              未填写 
                              <%else%>
                              <%=request.form("Shenfenzheng")%> 
                              <%end if%>
                            </td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#b4def8" align="right">联系电话：</td>
                            <td width="320" bgcolor="#b4def8">&nbsp; 
                              <%if request.form("Tel")="" then%>
                              未填写 
                              <%else%>
                              <%=request.form("Tel")%> 
                              <%end if%>
                            </td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#9FD6F8" align="right">详细地址：</td>
                            <td width="320" bgcolor="#9FD6F8">&nbsp; 
                              <%if request.form("Address")="" then%>
                              未填写 
                              <%else%>
                              <%=request.form("Address")%> 
                              <%end if%>
                            </td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#b4def8" align="right">邮政编码：</td>
                            <td width="320" bgcolor="#b4def8">&nbsp; 
                              <%if request.form("Youbian")="" then%>
                              未填写 
                              <%else%>
                              <%=request.form("Youbian")%> 
                              <%end if%>
                            </td>
                          </tr>
                          <tr> 
                            <td width="500" bgcolor="#67BDF1" align="center" colspan="2"> 
                              <input value="现在登陆" name=Login type=submit>
                            </td>
                          </tr>
                        </form>
                      </table>
                    </div>
                  </td>
                </tr>
                <tr> 
                  <td width="100%" height="12"></td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  
</div>       
<div align="center">        
  <center>
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


