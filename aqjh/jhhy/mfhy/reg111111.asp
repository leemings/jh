<%@ LANGUAGE=VBScript codepage ="936" %>
<%
if session("aqjh_jhdj")<20 then
	Response.Write "<script Language=Javascript>alert('提示：会员在线申请系统需要20级才可以！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"& session("aqjh_name") &"'",conn
if rs("会员等级")>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你现在已经是会员了,请到期后申请好不好？！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "SELECT * FROM q where d='未付' and  b<now()-70000 order by b DESC",conn,3,3
do while not rs.eof 
	conn.Execute "update 用户 set 状态='监禁',事件原因='监禁:\n您在我们这里办理会员："&rs("e")&"，\n因为你不付款，我们监禁你的账号！!' where 姓名='"&rs("a")&"'"
	conn.Execute "update q set d='监禁' where a='"&rs("a")&"'"
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
<title><%=Application("aqjh_chatroomname")%>用户注册</title>
<link rel="stylesheet" href="../css.css">
</head>
<body bgcolor="#8EB4D9">
<center>
<table border=1 bgcolor="#FFFFFF" align=center width=400 cellpadding="5" cellspacing="10" background="../images/b2.gif" height="227">
<tr bgcolor="#FFFFFF" align="center">
      <td height="314" bgcolor="#8EB4D9"> <font color="#000000">在这里你可以在线申请加入会员，
      购买会员</font>汇款办法请见<a href="../chat/hy.asp" target="_blank"><b><font color="#0000FF">这里</font></b></a>！<br>
        <br>
        <font color="#FF0000"> 办理会员是对我们江湖建设的最大支持！</font><br>
<br>
<table width="325" height="229">
<tr>
            <td height="236"> 
              <form method=POST action='regok.asp' onsubmit='return(check());' name=reg>
                <p align="center">申请人姓名：<font color=yellow><b><%=session("aqjh_name")%></b></font>
  请输入密码：    
                  <input type=password name=pass size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">   
                  <font color="#FF0000">*</font><br>   
                  <br>   
                  会员等级：    
                  <select name=hygrade onChange="javascript:js();" size=1 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">   
                    <option value="1" selected>一级</option>   
                  </select>   
                  级 申请时间：    
                  <select name=hymonth onChange="javascript:js();"  size=1 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">   
                    <option value="1" selected>01</option>   
                  </select>   
                  月<br>   
                  <br>   
                  所需金额：    
                  <input type=text name=money readonly size=5 maxlength="5" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="0">   
                  元 <br>   
                </p>   
<p align="center">   
<input type=submit value=申请 name="submit" style="border: 1px solid; font-size: 9pt; border-color:   
#000000 solid">   
<input type=button value=关闭 onClick="window.close()" name="button" style="border: 1px solid; font-size: 9pt; border-color:   
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
   
if(pass=="" || pass==null){window.alert("大侠，开玩了吧！不输入密码你申请什么会员啊？！");return false;}   
}   
</SCRIPT>