<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();</script>"
	Response.End
end if
%>
<html>
<head>
<title>�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
-->
</style>
<link rel="stylesheet" href="style.css" type="text/css">
</head>

<body bgcolor="#CC9900" text="#CC9900" style="font-size: 12px">
<div align="center">
<table border="0" width="86" height="338" cellspacing="0" cellpadding="0">
  <tr>
    <td width="136" height="278" valign="top" align="left"></td>
    <td width="219" height="278" valign="top" align="left">
      <div align="center">
      <table border="0" width="132" style="font-size: 15px" cellspacing="0" cellpadding="0" height="257" align="left">
        <tr> 
          <td height="14" width="130" style="border-left: 1 solid #FFCC66; border-right: 1 solid #FFCF63; border-top: 1 solid #FFCC66">
            <p align="center"><font color="#0000FF"><a href="yldok.asp">�����</a></font></p>
          </td>
        </tr>
        <tr>
          <td height="141" width="130" style="border-left: 1 solid #FFCC66; border-right: 1 solid #FFCF63; border-top: 1 solid #FFCC66; border-bottom: 1 solid #FFCF63" align="center" valign="top">��<p>
			<font color="#800000">ʹ���������Ҫ</font><font color="#0000FF">ת��3�Σ�5000���ֽ�,10�����,20����</font><font color="#000080">,</font><font color="#008000">���ӹ���10000����10000,����������������100��200��</font><font color="#000080">,</font><font color="#0000FF">ע�⹥�������������޵�,�ಹ��û��ʲô��˼!~</font>~&nbsp;<font size="2">&nbsp; </font>&nbsp;</td>
        </tr>
      </table>
      </div>
    </td>
  </tr>
</table></div>
</body>
</html>
