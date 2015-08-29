<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="md5.asp"-->
<!--#include file="inc/chkinput.asp"-->
<%
'=========================================================
' File: modifyadd.asp
' Version:5.0
' Date: 2002-9-11
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

stats="修改联系资料"
dim sex
dim showre
dim password
dim quesion
dim answer
dim face,width,height
dim email
dim birthday
dim usercookies

call nav()
if Cint(GroupSetting(16))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛修改资料的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	founderr=true
end if
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您还没有登录，请登录后进行修改。"
	founderr=true
end if

if founderr then
  	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call head_var(0,0,membername & "的控制面板","usermanager.asp")
	if request("action")="updat" then
		if lcase(membername)="test" then
		  errmsg=errmsg+"<br>"+"<li>测试用户不能修改个人资料。"
			call dvbbs_error()
		else
			call update()
			if founderr then
				call dvbbs_error()
			else
				sucmsg="更新联系资料成功！"
				call dvbbs_suc()
			end if
		end if
	else
		call userinfo()
		if founderr then call dvbbs_error()
	end if
end if
call activeonline()
call footer()

if request("action")="updat" then

end if

sub userinfo()
set rs=server.createobject("adodb.recordset")
sql="Select * from [User] where userid="&userid
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>该用户名不存在。"
	founderr=true
	exit sub
else
%>
<!--#include file="z_controlpanel.asp"-->
<br>
<form action="modifyadd.asp?action=updat&username=<%=htmlencode(membername)%>" method=POST name="theForm">
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr> 
      <th colspan="2" width="100%">用户联系资料
      </th>
    </tr>
      <tr>    
        <td width="40%" class=tablebody1><B>主页地址</B>：<BR>如果您有主页，请输入主页地址。此项可选</td>   
        <td width="60%" class=tablebody1>    
          <input type="TEXT" name="homepage" value="<%if trim(rs("homepage"))<>"" then%><%=htmlencode(rs("homepage"))%><%end if%>" size=30 maxlength=100>   
        </td>   
      </tr>  
      <tr>    
        <td width="40%" class=tablebody1><B>Email地址</B>：<BR>请输入有效的邮件地址，这将保证您在论坛中的私人资料。</td>   
        <td width="60%" class=tablebody1>    
          <input type="TEXT" name="Email" size=30 maxlength=50 value="<%if trim(rs("useremail"))<>"" then%><%=htmlencode(rs("useremail"))%><%end if%>">   
        </td>   
      </tr>    
      <tr>    
        <td width="40%" class=tablebody1><B>OICQ号码</B>：<BR>如果您有 OICQ，请输入号码。此项可选</td>   
        <td width="60%" class=tablebody1>    
          <input type="TEXT" name="OICQ" value="<%if rs("oicq")<>"" then%><%=htmlencode(rs("oicq"))%><%end if%>" size=30 maxlength=20>   
        </td>   
      </tr>
      <tr>    
        <td width="40%" class=tablebody1><B>ICQ 号码</B>：<BR>如果您有 ICQ，请输入号码。此项可选</td>   
        <td width="60%" class=tablebody1>    
          <input type="TEXT" name="ICQ" value="<%if rs("icq")<>"" then%><%=htmlencode(rs("icq"))%><%end if%>" size=30 maxlength=20>   
        </td>   
      </tr>
      <tr>    
        <td width="40%" class=tablebody1><B>MSN 号码</B>：<BR>如果您有 MSN，请输入。此项可选</td>   
        <td width="60%" class=tablebody1>    
          <input type="TEXT" name="msn" value="<%if rs("msn")<>"" then%><%=htmlencode(rs("msn"))%><%end if%>" size=30 maxlength=50>   
        </td>   
      </tr>
    <tr align="center"> 
      <td colspan="2" width="100%"  class=tablebody2>
            <input type=Submit value="更 新" name="Submit"> &nbsp; <input type="reset" name="Submit2" value="清 除">
      </td>
    </tr>

    </table></td></tr></table></form>
   
</body>
</html> 
<%
end if
rs.close
set rs=nothing
end sub

sub update()
if IsValidEmail(request("Email"))=false then
	errmsg=errmsg+"<br>"+"<li>您的Email有错误。"
	founderr=true
	exit sub
else
	if not isnull(forum_setting(52)) and forum_setting(52)<>"" and forum_setting(52)<>"0" then
		Dim SplitUserEmail
		SplitUserEmail=split(forum_setting(52),"|")
		For i=0 to ubound(SplitUserEmail)
			if instr(request("email"),SplitUserEmail(i))>0 then
			errmsg=errmsg+"<br>"+"<li>您填写的Email地址含有系统禁止字符。"
   			founderr=true
			exit sub
			end if
		Next
	end if
	Email=request("Email")
end if
if request("oicq")<>"" then
	if not isnumeric(request("oicq")) or len(request("oicq"))>12 then
		errmsg=errmsg+"<br>"+"<li>Oicq号码只能是4-12位数字，您可以选择不输入。"
		founderr=true
		exit sub
	end if
end if
set rs=server.createobject("adodb.recordset")
sql="Select * from [User] where userid="&userid
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>该用户名不存在。"
	founderr=true
	exit sub
else
	rs("useremail")=email
	rs("homepage")=checkStr(trim(request.form("homepage")))
	rs("oicq")=request.form("oicq")
	rs("icq")=request.form("icq")
	rs("msn")=checkStr(trim(request.form("msn")))
      
	rs.Update
end if
rs.close
set rs=nothing
end sub
%>