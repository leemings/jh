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
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>��̥���ǣ��õ�����</TITLE>
<BGSOUND src="image/baihe.mid" loop=-1>
<link rel="stylesheet" type="text/css">
<style>
<!--
.6     { background-color: #E1E9FF; vertical-align: 1 }
-->
</style>
</HEAD>
<BODY vLink=#e1e9ff aLink=#e1e9ff link=#e1e9ff bgColor=#78869f background="image/wonderofitallbkgtile2a.jpg" style="font-size: 8pt; color: #FFFFFF">
<DIV align=center>
<CENTER>
<TABLE cellSpacing=20 borderColorDark=#3d4b61 cellPadding=0 width="88%" borderColorLight=#8f9fba background="image/wonderofitallstrip.jpg" border=1>
<TBODY> 
<TR>
<TD width="100%" align=center>
<TABLE cellSpacing=10 borderColorDark=#3d4b61 cellPadding=0 width="100%" borderColorLight=#8f9fba background="image/wonderofitalltilelight.jpg" border=1>
 <TBODY> 
 <TR>
 <TD width="100%" height="411" align=center> 
<TABLE cellPadding=0 width="100%" background="image/wonderofitallstrip.jpg" border=0>
 <TBODY>
 <TR>
 <TD width="100%" align=center>
<TABLE cellPadding=0 width="100%" background="image/wonderofitalltile.jpg" border=0>
<TBODY>
 <TR align=center>
<TD width="100%" height="372" valign="top"> 
<table width="594" border="0" cellspacing="0" cellpadding="0" height="25">
<tr> 
<td width="235" height="31" valign="middle" align="center"><br>
<br>
                                          <font color="FFFFFF" zize=+3"><b>����̥���ǣ��õ�������</font> 
                                        </td>
<td width="359" height="235" rowspan="2"> <img height=274 src="image/right.jpg" width=360 border=0></td>
</tr>
<tr>
<td width="235" height="215" valign="top">
<br><br>
                                          <font size="2"><font color="#FFFF00"><b>˵&nbsp;&nbsp; 
                                          ����</font> �����ȼ�����<font color="#FFFF00"><b>200��</font>����������õ������������õ���������ĵȼ�����Ϊ<font color="#FFFF00"><b>0��</font>���õ���������������������书���޶����кܴ����ߣ��������������ļ���Ϊ: 
                                          <font color="#FFFF00"><b>�������������������������ݵ��ܻ��ֵ�ʮ��֮һ��1/10�����书�����������������ݵ��ܻ��ֵ�ʮ��֮һ��1/10��</font>����<br>
                                          �ر������������ͷ���ֻ���ȼ��йأ����Եõ���������Ĺ�����������Ϊ1��ʱ����ֵ������װ���������ӹ�������
</font>
<br>
<br>
<br>
<div align="center"><a href="zsok.asp"><font color="#FFFF00" size=+2><b>��Ҫ��̥���ǣ��õ�����</font> 
                                          </div>
</td></form>
</tr>
</table>
<a href="#" onclick="javascript:window.close();"><img height=53 src="image/wonderofitallbutterflydark.jpg" width=60 border=0 alt="�رմ���"></a><br>
<IMG height=25 src="image/wonderofitallbar.jpg" width=488 border=0> 
<br>
                            <font size="2" color="#FFFFFF">&copy; Copyright 2005 
                            www.7758530.com &#153; All Rights Reserved ���ֽ��� 
                            ����Ȩ���� ��ֹ��Ϯ<br>
                            ������������׵��ꡡQQ:865240608</font> </TD>
                        </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
</TD></TR></TBODY></TABLE></CENTER></DIV></BODY></HTML>
