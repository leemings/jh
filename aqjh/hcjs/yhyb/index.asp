<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<HTML><HEAD><TITLE>������</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../../css.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#fffddf leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" background="../../bg.gif">
<div align="center">
<table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#000000" width="32%" height="31">
<tr align="center"> 
<td bgcolor="#669900" width="100%" height="25"><b><font size="4" color="#FFFFFF">������<img border="0" src="../../chat/img/tu1.gif"></font></b></td>
</tr>
</table>
<br>
</div>
<table border=0 cellpadding=0 cellspacing=0 width="601" align="center">
<tbody> 
<tr> 
<td class=bg1 colspan=2 rowspan=2 
style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" valign=top> 
<table border="0" cellpadding="0" cellspacing="0" width="554">
<tr> 
          <td align="center" colspan="2" width="552"> <b> 
 <font size="2" color="#65251C"> 
            <font color="#FF0000"><%=aqjh_name%></font></font><font size="2" color="#008080">���������ݣ��������������ҵ���������Ѽ������Ŷ</font> 
            </b> 
          </td>
</tr>
<tr> 
<td align="center" rowspan="2" width="249"><img border="0" src="../../chat/img/004.gif"><img border="0" src="../../chat/img/17s.jpg"> 
</td>
          <td width="301"> <font size="3" color="#65251C"><b><font color="#FF0000">�ƹ��޻�˵����</font></b></font><font size="2" color="#800000">��Ϻ���쵽�������������������Ժ�����������Ƿ�йร�������������Ŷ��������������������������ղ�����Ŷ������������У�ûǮ�Ļ�����Ҳ���Ե�������������Ŷ��������������ò�����Ļ��Ͳ��������ˣ�������������ֻҪ����Ŷ��</font> 
          </td>
</tr>
<tr> 
<td width="301" valign="baseline" align="center"> 
<p><b><a href="xiaojie.asp"><font size="2" color="#669900">�������֣�û��Ҫ��̩���ɣ�����������������</font> 
</a></b> 
</p>
<p><font size="2" color="#65251C"><b><a href="dengji.asp">����ûǮ�ˣ���Ҫ�����ϰࡣ</a><br>
</b></font></p>
<p><b><font size="2"><a href="delgirl.asp"><font color="#FF0000">�����ˣ���Ҫ��������</font></a></font></b></p>
<p>&nbsp; </p>
</td>
</tr>
</table>
</td>
</tr>
<tr> 
<td colspan=2>&nbsp;</td>
</tr></tbody></table></BODY></HTML> 