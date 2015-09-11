<%@ LANGUAGE=VBScript codepage ="936" %>
<%
if session("aqjh_jhdj")<250 or session("aqjh_jhdj")>260 then
	Response.Write "<script Language=Javascript>alert('提示：在线申请免费一级会员需要250级才可以！超过260级视为自动放弃,请申请更高的奖励另一奖励');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
'rs.open "select * from 用户 where 姓名='"& session("aqjh_name") &"'",conn
'if rs("会员等级")>0 then
'	rs.close
'	set rs=nothing
'	conn.close
'	set conn=nothing
'	Response.Write "<script Language=Javascript>alert('提示：你现在已经是会员了请与站长联系\n办理会员续费！！');location.href = 'javascript:history.go(-1)';</script>"
'	Response.End
'end if
'rs.close
rs.open "SELECT * FROM q where d='未``付' and  b<now()-100000 order by b DESC",conn,3,3
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
<title>爱情江湖总站免费会员申请系统-----免费2级会员申请</title>
<link rel="stylesheet" href="../css.css">
</head>
<body bgcolor="#000080" background="BJ.jpg">
<center>
<table border=1 bgcolor="#FFFFFF" align=center width=400 cellpadding="5" cellspacing="10" background="BJ.jpg" height="227" bordercolorlight="#008000" bordercolordark="#00FFFF" bordercolor="#FFFF00">
<tr bgcolor="#FFFFFF" align="center">
      <td height="314" bgcolor="#8EB4D9"> <font color="#000000">在这里你可以在线申请加入免费会员，在申请后请您将自动获得免费二级会员一个月，如果您想得到更多的奖励，请为本江湖多多拉人，泡点…………</font><b><font color="#FFFF00">加入付费会员享受更高贵服务</font></b>&nbsp; 
        請點<a href="../chat/hy.asp" target="_blank"><b><font color="#0000FF">这里</font></b></a>！<br>
        <br>
        <font color="#FF00FF"><b><span style="background-color: #8EB4D9"><font face="隶书" size="4">办理会员是对我们江湖建设最大支持</font></span></b>！<br>
        </font>
        <img border="0" src="gg01.jpg" width="213" height="60">
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
                    <option value="2" selected>二级</option>
                    
                  </select>
                  级 申请时间： 
                  <select name=hymonth onChange="javascript:js();"  size=1 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                    <option value="1" selected>01</option>  
                  <br>
                  
                </p>
<p align="center">
<input type=submit value=申请 name="submit" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
<input type=button value=关闭 onClick="window.close()" name="button" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
</p>
</form>
              <p align="center"><font face="方正舒体" size="5"><b><span style="background-color: #00FF00">爱情江湖,創意無限</span></b></font><font color="#000000"><%=Application("aqjh_chatroomname")%></font><font size="2"><br>
                </font></p></td>
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

if(pass=="" || pass==null){window.alert("大侠，不输入密码是不可以申请会员的！");return false;}
}
</SCRIPT>