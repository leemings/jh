<!--#include file="conn.asp" -->
<!--#include file="z_Tigerset.asp" -->

<%dim membername,username,server_v1,server_v2,BSWIN,BSm,rs,nowcash
membername=request.cookies("aspsky")("username")
username=CheckLogin
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then%>
	<script language="JavaScript">
		alert("你提交的路径有误，禁止从站点外部提交数据请不要乱该参数");
	</script>
	<script language="VBScript">
		top.location.replace("z_tiger.asp")
	</script>
	<%response.end
end if

If Request.form("Submit")<>"" then
conn.Execute("Update ["&DatabaseMember&"] set "&DatabaseCash&"="&DatabaseCash&"+"&Request.form("win")&" where "&DatabaseUserName&"='"&username&"'")
%>
<script language="VBScript">
	top.location.replace("z_tiger.asp")
</script>
<% End if

BSWIN=Request.form("BSWIN")
BSm=Request.form("BSm")

set rs=Createobject("ADODB.Recordset")
rs.open"select * from ["&DatabaseMember&"] where "&DatabaseUserName&"='"&username&"'",conn,1,1
nowCash=clng(rs(DatabaseCash))
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css" type="text/css">
<style type="text/css">

P {
	FONT-SIZE: 9pt; LINE-HEIGHT: 12pt
}
A {
	TEXT-DECORATION: none;
	TEXT-TRANSFORM: none;
	color: #000000;

}
A:hover	{color:#0099FF; text-decoration: underline}
TD {
	FONT-FAMILY: 宋体,simsun; FONT-SIZE: 9pt; line-height: 130%
}
DIV {
	FONT-FAMILY: 宋体,simsun; FONT-SIZE: 9pt
}
INPUT {
	BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000; FONT-FAMILY: 宋体, sans-serif, serif; FONT-SIZE: 9pt
}
SELECT {
	BORDER-BOTTOM: 1px double; BORDER-LEFT: 0px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000; FONT-FAMILY: 宋体, sans-serif, serif; FONT-SIZE: 9pt
}
SELECT:unknown {
	BORDER-BOTTOM: 0px outset; BORDER-LEFT: 0px outset; BORDER-RIGHT: 0px outset; BORDER-TOP: 0px outset; CURSOR: hand; FONT-FAMILY: 宋体, sans-serif, serif; FONT-SIZE: 9pt
}
BLOCKQUOTE {	
}
</style>
<script language="VBScript">

dim n,o,m
sub bs(mm)
	form2.Submit.disabled=true
	o=1
	randomize
	n=fix(11*rnd+50)
	if mm=1 then
		tob()
	else
		tos()
	end if
end sub

sub tob()
	m=m+1
	if m>9 then m=0
	cc2.innerHTML=m
	o=o+1
	if o>n then
		if m>=5 then
			form1.BSWIN.value=form1.BSWIN.value*2
		else
			form1.BSWIN.value=0
		end if
		form1.BSm.value=m
		form1.submit()
		form2.Submit.disabled=false
		exit sub
	end if
	settimeout "tob()",20
end sub

sub tos()
	m=m+1
	if m>9 then m=0
	cc2.innerHTML=m
	o=o+1
	if o>n then
		if m<=4 then
			form1.BSWIN.value=form1.BSWIN.value*2
		else
			form1.BSWIN.value=0
		end if
		form1.BSm.value=m
		form1.submit()
		form2.Submit.disabled=false
		exit sub
	end if
	settimeout "tos()",20
end sub
</script>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#99CCFF">
<% If BSWIN>0 then %>
<table width="254" border="0" cellspacing="0" cellpadding="0" height="180">
  <tr>			
    <td> 
      <table width="254" border="0" cellspacing="0" cellpadding="3">
        <tr> 
          <td>您目前有现金：<%=nowCash%></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td align="center" height="141"> 
			<table width="240" border="0" cellspacing="0" cellpadding="0">
			  <tr align="center"><td colspan="3" height="30">WIN：<%=BSWIN%></td></tr>
			  <tr align="center">
					<td width="80" height="80"><img src="duchang_img/Big.gif" width="46" height="40" onClick="bs(1)" style="cursor:hand"></td>
					<td width="80" height="80" id=cc2><%=BSm%></td>
					<td width="80" height="80"><img src="duchang_img/Small.gif" width="46" height="40" onClick="bs(2)" style="cursor:hand"></td>
			  </tr>
			  <tr align="center"><form name="form2" method="post" action="z_TigerWin.asp"><td colspan="3" height="30"><input type="submit" name="Submit" value="下一盘" class="bt"><input type="hidden" name="win" value="<%=BSWIN%>"></td></form></tr>
			</table>
    </td>
  </tr>
	<form name="form1" method="post" action="z_TigerWin.asp"><input type="hidden" name="BSWIN" value="<%=BSWIN%>"><input type="hidden" name="BSm" value=""></form>
</table>
<% ElseIf BSWIN=0 then %>
<table width="254" border="0" cellspacing="0" cellpadding="0" height="180">
  <tr>			
    <td height="60"> 
      <table width="254" border="0" cellspacing="0" cellpadding="3">
        <tr><td>您目前有现金：<%=nowCash%></td></tr>
      </table>
    </td>
  </tr>
  <tr><td align="center" height="141"><p><font color="#FF0000">押大小没押中！</font></p><p><input type="button" name="Submit2" value="下一盘" class="bt" onclick="top.location.replace('z_tiger.asp')"></p></td></tr>
</table>
<% End if %>
</body>
</html>
