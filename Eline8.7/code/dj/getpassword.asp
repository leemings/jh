<!--#include file="conn.asp"-->
<title>用户取回密码</title>
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
                  <b>取回密码</b>（第一步：输入您的用户名）</td>
          </tr>
                <tr> 
                  <td height="126">
                  <p align="center"><font color="#FF0000">请输入您注册时的用户名:<br>
                  <INPUT name=username type=text size=30 value="在这里输入您的用户名" onclick="Javascript:this.value=''"></font></td>
                  </tr>
    	        <tr> 
                  <td align=center height="1"> 
                    <input type=submit name="submit" value="提交，进入下一步">
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
response.write "<script language=javascript>alert('请输入用户名！！！');history.back(1);</script>"
	founderr=true
else
set rs=server.createobject("adodb.recordset")
	sql="select username,quesion from User where username='"&username&"'"
rs.open sql,conn,1,3
	if rs.bof and rs.eof then
response.write "<script language=javascript>alert('该用户名并没有在本站注册过，请确认输入是否正确！！！');history.back(1);</script>"
	founderr=true
	elseif rs("quesion")="" or isnull(rs("quesion")) then
response.write "<script language=javascript>alert('操作失败！你注册时没有填写取回密码时所需要的信息！！！');history.back(1);</script>"
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
                  <b>取回密码</b>（第二步：回答问题）</td>
          </tr>
                <tr> 
                  <td  valign=middle width="121" height="67"> <b>密码提示问题:</b></td>
                  <td width="179" height="67"><font color="ff0000"><%=rs("quesion")%></font>　</td>
                </tr>
	            <tr> 
                  <td width="121" height="1"><b>取回密码答案:</b></td>
                  <td width="179" height="1"> 
                    <INPUT name=answer type=text size=30></td></tr>
	            <tr> 
                  <td colspan=2 align=center height="87"> 
                    <input type=submit name="submit" value="提交，答案已经填写正确">
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
response.write "<script language=javascript>alert('请输入问题的答案！！！');history.back(1);</script>"
	founderr=true
else
	sql="select answer,PassWord from user where username='"&username&"'"
	set rs=conn.execute(sql)
	if rs("answer")<>answer then
response.write "<script language=javascript>alert('操作失败！你的问题回答得不正确！！！');history.back(1);</script>"
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
                  <b>取回密码</b>（成功取回）</td>
    </tr>
                <tr> 
                  <td height="12" width="400">
                  <p align="center">&nbsp;&nbsp;恭喜您，您成功取回密码！</td>
          </tr>
	            <tr> 
                  <td height="58" width="400" align="right">
                  <p align="center"><b>用户名:<font color="#ff0000"><%=username%></font></b></p>
                  <p align="center"></p>
                  <p align="center">
                  <b>密&nbsp;&nbsp;码:<font color="#ff0000"><%=rs("PassWord")%></font></b></td>
                </tr>
	            <tr> 
                  <td height="35" width="400">
                  <p align="center"><b>说&nbsp;&nbsp;明:</b>请妥善保管您的密码！</td>
                </tr>
                <tr align="center"> 
                  <td height="36" width="400"> 
                    <input type=button name=login value=成功取回密码，关闭此窗口 onclick=window.close()>
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