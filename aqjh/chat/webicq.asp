<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0%>
<%

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Session("aqjh_inthechat")<>"1" then 
	Response.Write "<script Language=Javascript>alert('�����������ٺ��У�');window.close();</script>"
	Response.End
end if
if aqjh_jhdj<10 then 
	Response.Write "<script Language=Javascript>alert('ս���ȼ�����10�����ɺ��У�');window.close();</script>"
	Response.End
end if

useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
aqjh_name=aqjh_name
who=Trim(Request.Form("who"))
if who="" then who=aqjh_name
show=Split(Trim(useronlinename)," ",-1)
x=UBound(show)
chatroombgimage=Application("aqjh_chatimage")
chatroombgcolor=Session("afa_chatbgcolor")%><html>
<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
<script Language=JavaScript>
function check(){if(document.forms[0].intro.value==""){alert("��Ϣ���ݲ���Ϊ�գ�");return false;}document.forms[0].Submit.disabled=1;return true;}
</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=<%=chatroombgcolor%> background=<%=chatroombgimage%> leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr>
<td>
<table width=350 border=1 align=center cellspacing=0 bordercolorlight=000000 bordercolordark=FFFFFF bgcolor=E0E0E0>
<tr valign="top">
<td>
<table border=0 bgcolor=#3399FF cellspacing=0 cellpadding=2 width=344>
<tr>
<td width=326><font color=FFFFFF face=Wingdings>*</font><font color=FFFFFF>���� Web ICQ ��Ϣ</font></td>
<td width=18>
<table border=1 bordercolorlight=666666 bordercolordark=FFFFFF cellpadding=0 bgcolor=E0E0E0 cellspacing=0 width=18>
<tr>
<td width=16><b><a href="javascript:window.close()" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="�˳�"><font color="000000">��</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table width="100%" border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="FFFFFF" cellpadding="0">
<tr valign="middle" align="center">
<td class=p9>
<table width="100%" border="0" cellpadding="3">
<tr align="center">
<td width="25%" bgcolor="#000000"><font color="FFFFFF">������Ϣ</font></td>
<td width="25%" bgcolor="E0E0E0"><a href=webicqlist.asp><font color="black">��Ϣ�б�</font></a></td>
<td width="25%">&nbsp;</td>
<td width="25%">&nbsp;</td>
</tr>
</table>
<table width="100%" border="1" height="200" cellspacing="0" cellpadding="5" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="E0E0E0">
<form method="post" action="webicqsend.asp" onsubmit='return(check());'>
<tr>
<td>
<div align="center"><font color="#FF0000">��Ϣ�ĳ��ȱ���С��1024�ַ�<br>
</font><br>
</div>
<table border="0" align="center">
<tr>
<td>����</td>
<td>
<select name="towho">
<%for i=0 to x%>
<%if instr(Application("aqjh_admin"),show(i))=0 and show(i)<>aqjh_name then%>

<option value="<%=show(i)%>"<%if CStr(show(i))=CStr(who) then Response.Write " selected"%>><%=show(i)%></option>
<%end if%>
<%next%></select>���� <font color=red><%=x+1%></font> �� (ALT+S=����)
</td>
</tr>
<tr>
<td>��Ϣ��</td>
<td>
<textarea name="intro" cols="40" rows="5"></textarea>
</td>
</tr>
<tr>
<td colspan="2" align="center">
<input type="submit" name="Submit" value="����" accesskey='s'>
<input type="reset" value="��д">
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
</html>