<!--#include file="conn.asp" -->
<!--#include file="z_Tigerset.asp" -->
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
	FONT-FAMILY: ����,simsun; FONT-SIZE: 9pt; line-height: 130%
}
DIV {
	FONT-FAMILY: ����,simsun; FONT-SIZE: 9pt
}
INPUT {
	BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000; FONT-FAMILY: ����, sans-serif, serif; FONT-SIZE: 9pt
}
SELECT {
	BORDER-BOTTOM: 1px double; BORDER-LEFT: 0px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000; FONT-FAMILY: ����, sans-serif, serif; FONT-SIZE: 9pt
}
SELECT:unknown {
	BORDER-BOTTOM: 0px outset; BORDER-LEFT: 0px outset; BORDER-RIGHT: 0px outset; BORDER-TOP: 0px outset; CURSOR: hand; FONT-FAMILY: ����, sans-serif, serif; FONT-SIZE: 9pt
}
BLOCKQUOTE {	
}
</style>
<%
dim membername,username,rs,sql,nowcash,etotal
membername=request.cookies("aspsky")("username")
username=CheckLogin
	set rs=server.createobject("adodb.recordset")
	sql="select * from ["&DatabaseMember&"] where "&DatabaseUserName&"='"&username&"'"
	rs.open sql,conn,1,1
nowCash=clng(rs(DatabaseCash))
If eTotal>nowCash then%>
	<script language="JavaScript">
		alert("����Ǯ����");
	</script>
	<script language="VBScript">
		top.location.replace("z_tiger.asp")
	</script>
	<%response.end
End if
rs.close
set rs=nothing
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css" type="text/css">
<script language="VBScript">

sub help()
msgbox"�����Ӧˮ����ע���������ֿ�ʼ��"&chr(10)&"Ѻ�к�����ֱ�ӵ����һ�̱��潱��"&chr(10)&"Ҳ���ٴ�ͨ���´�Сʹ���Ľ��𷭱���",0,"�����ϻ�����Ϸ����"
end sub
</script>
</head>

<body bgcolor="#99CCFF" leftmargin="0" topmargin="0">
<table width="258" border="0" cellspacing="0" cellpadding="0" height="180">
  <tr> 
    <td width="254"> 
      <table width="254" border="0" cellspacing="0" cellpadding="3">
        <tr> 
          <td>��Ŀǰ���ֽ�<%=nowCash%>Ԫ</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td align="center" height="90" width="254">���ڵȴ���ʼ</td>
  </tr>
  <tr>
    <td align="center" height="51"><a href="#" onclick="help()">��Ϸ����</a></td>
  </tr>
</table>
</body>
</html>