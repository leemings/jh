<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
if aqjh_name="" or Session("aqjh_inthechat")<>"1" or Instr(useronlinename," "&aqjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"%><html>
<script Language="JavaScript">
function setfs(){parent.f2.document.af.fs.value=document.forms[0].ftsz.value;parent.f2.document.af.lh.value=document.forms[0].lheight.value;this.location.href="setfontsizeok.asp";}
</script>
<head>
<title>�����ֺż��о�</title>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<div align=center>
<table width="100%" border="0" height="100%">
<tr>
<td>
<table border="1" bgcolor="E0E0E0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" cellpadding="4">
<form method="post" action="javascript:setfs()" name="">
<tr>
<td>
<div align=center>
<p><font color="#FF0000" style="font-size:12pt">�ֺż��о�</font></p>
<p>��ѡ��������Ҫ�������Сֵ���Ի������о�ֵ��Ȼ�������ύ��</p>
<p>
<select name="ftsz">
<option value="9">9 ��</option>
<option value="9.5">9.5 ��</option>
<option value="10">10 ��</option>
<option value="10.5">10.5 ��</option>
<option value="11">11 ��</option>
<option value="11.5">11.5 ��</option>
<option value="12">12 ��</option>
<option value="12.5">12.5 ��</option>
<option value="13">13 ��</option>
<option value="13.5">13.5 ��</option>
<option value="14">14 ��</option>
<option value="14.5">14.5 ��</option>
<option value="15">15 ��</option>
<option value="15.5">15.5 ��</option>
<option value="16">16 ��</option>
</select></p>
<p>
<select name="lheight">
<option value="120">������%</option>
<option value="125">������%</option>
<option value="130">������%</option>
<option value="135">������%</option>
<option value="140">������%</option>
<option value="145">������%</option>
<option value="150">������%</option>
<option value="155">������%</option>
<option value="160">������%</option>
<option value="165">������%</option>
<option value="170">������%</option>
<option value="175">������%</option>
<option value="180">������%</option>
<option value="185">������%</option>
<option value="190">������%</option>
<option value="195">������%</option>
<option value="200">������%</option>
<option value="205">������%</option>
<option value="210">������%</option>
<option value="215">������%</option>
<option value="220">������%</option>
<option value="225">������%</option>
<option value="230">������%</option>
<option value="235">������%</option>
<option value="240">������%</option>
<option value="245">������%</option>
</select>
</p>
<p>
<input type="submit" name="Submit" value="�ύ">
</p>
</div>
</td>
</tr>
</form>
</table>
</td>
</tr>
</table>
</div>
<Script Language=Javascript>
document.forms[0].ftsz.value=parent.f2.document.af.fs.value;
document.forms[0].lheight.value=parent.f2.document.af.lh.value;
parent.m.location.href='about:blank';
</Script>
</body>
</html>