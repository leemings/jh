<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if session("sjjh_name")="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ��������ѵ�����Ƚ��뽭�������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<!--#include file="conn.asp"-->
<!--#include file="function.asp"-->
<%
id=request("id")
if id="" then
	errmsg="����δѡ�����!"
	call error()
	Response.End 
end if
%>
<%
set rs=conn.execute("SELECT * FROM MusicList where id="+id)
if rs.eof then
	errmsg="����ѡ���������ѡ!"
	call error()
	Response.End 
else
	MusicName=rs("MusicName")
	Singer=rs("Singer")
	Wma=rs("Wma")
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��裺<%=Singer%>--{<%=MusicName%>}, </title>
<link rel=stylesheet href=images/style.css>
</head>
<body bgcolor="#B4DEF8" topmargin="0" leftmargin="0">             
<table border="1" width="100%" cellspacing="0" bordercolor="#56B0F4" bordercolordark="#56B0F4"> 
  <tr> 
    <td width="100%" height=22 Bgcolor="#56B0F4" align="center"><b>EMAIL���͵�赥</b></td>
  </tr>
  <tr>
    <td width="100%">               
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <form name="form1" method="post" action="Sendmail.asp">   
          <tr>
            <td width=50% bgcolor="#B4DEF8">&nbsp;&nbsp; <b>�͸�˭</b>��<INPUT type=text name=name size=20 value='�����ѵ�����' onFocus=this.value=''></td>
            <td width=50% bgcolor="#B4DEF8">&nbsp;<b>Email</b>�� <INPUT type=text name=mail size=20 value='�����ѵ�����' onFocus=this.value=''></td>
          </tr>
          <tr>
            <td width=50% bgcolor="#B4DEF8">&nbsp;&nbsp; <b>�����</b>��<INPUT type=text name=cname size=20 value='��������' onFocus=this.value=''></td>
            <td width=50% bgcolor="#B4DEF8">&nbsp;<b>Email</b>�� <INPUT type=text name=cmail size=20 value='��������' onFocus=this.value=''></td>
          </tr>
          <tr>
            <td width=100% valign=top colspan="2" bgcolor="#B4DEF8" align="center"><TEXTAREA name=dmail rows=4 cols="60">ף��...</TEXTAREA></td>
          </tr>
          <tr>
            <td width=100% height=23 bgcolor="#56B0F4" align=center colspan="2"><input type="submit"  value="�ύ" name="Submit"> 
              &nbsp;&nbsp; <input type="reset"  value="��д" name="B1"></td>
          </tr>			
      </table><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
<input type="text" name="email" value="<%=Singer%>--{<%=MusicName%>}" class=kuang size="42" maxlength="50">
      <input type="text" name="fmail" value="http://zhzx.jjedu.org/eline/music/MusicPlay.asp?id=<%=id%>"��class=kuang size="42" maxlength="50">
    </td>
  </FORM></tr>
</table>
</body>
</html>