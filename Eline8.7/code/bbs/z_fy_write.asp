<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="z_fy_const.asp"-->
<%
'=========================================================
' Plusname: ��Ժ����������
' File: z_fy_write.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
Response.Buffer=True 
stats="׫д״ֽ"
call nav()
call head_var(0,0,"������Ժ","z_fy_fayuan.asp")
session("username")=membername
if not founduser then
Errmsg=Errmsg+"<br>"+"<li>��û���ڱ�������Ժ��Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
call dvbbs_error()
else
  response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=25><b>׫д״��</b></td></tr><tr><td valign=middle class=tablebody2 height=100 ><CENTER>"
  call getfyconfig()
  response.write "<form action=z_fy_newplan.asp method=post  width=60% class=tablebody1>"
 response.write "<table class=tableborder1 cellspacing=1 cellpadding=3 align=center >"
 response.write "<tr><td class=tablebody2 rowspan=3 width=10% > ע�����<br><br>��Ͷ������Ҫ�������Ҫ�����������20�������ڡ�<br>�ڶԱ���Ĵ���Ҫ���������������ȷ������ע��߶ȡ�<br>��Ͷ�����ݺ���������ϸ��д���֤�ݣ�����漰��̳����Υ����������������λ�á�</td>"
 response.write "<td colspan=2 width=20% class=tablebody1><b>Ͷ�����⣺</b><input name=topic size=86 maxlength=86></td></tr>"
 response.write "<tr><td width=15%  class=tablebody1><b>�������棺</b><input name=towho size=22 maxlength=22></td>"
 response.write "<td width=20%  class=tablebody1><b>Ҫ��Ա�����У�</b>" 
 response.write "<select name=play size=1>"
 response.write "<option value='�������'><b>���������̡�</b></option>"
 response.write "<option value='����10����'><b>&nbsp;&nbsp;������10����</b></option>"
 response.write "<option value='������Сʱ'><b>&nbsp;&nbsp;��������Сʱ</b></option>"
 response.write "<option value='����һСʱ'><b>&nbsp;&nbsp;������һСʱ</b></option>"
 response.write "<option value='����һ��'><b>&nbsp;&nbsp;������һ��</b></option>"
 response.write "<option value='��������'><b>&nbsp;&nbsp;����������</b></option>"
 response.write "<option value='����һ��'><b>&nbsp;&nbsp;������һ��</b></option>"
 response.write "<option value='û��ȫ���Ʋ�'><b>���û��ȫ���Ʋ���</b></option>" 
 response.write "<option value='����1��'><b>&nbsp;&nbsp;������1%</b></option>"
 response.write "<option value='����10��'><b>&nbsp;&nbsp;������10%</b></option>"
 response.write "<option value='����50��'><b>&nbsp;&nbsp;������50%</b></option>"
 response.write "<option value='�۳����о���'><b>��Կ۳����о����</b></option>"               
 response.write "<option value='�۾���1��'><b>&nbsp;&nbsp;���۾���1%</b></option>"
 response.write "<option value='�۾���10��'><b>&nbsp;&nbsp;���۾���10%</b></option>"               
 response.write "<option value='�۾���50��'><b>&nbsp;&nbsp;���۾���50%</b></option>"
 response.write "<option value='�۳���������'><b>��Կ۳�����������</b></option>" 
 response.write "<option value='������1��'><b>&nbsp;&nbsp;��������1%</b></option>"               
 response.write "<option value='������10��'><b>&nbsp;&nbsp;��������10%</b></option>"               
 response.write "<option value='������50��'><b>&nbsp;&nbsp;��������50%</b></option>"
 response.write "<option value='���'><b>������з�ֵ�����</b></option>"
 response.write "<option value='����-10'><b>&nbsp;&nbsp;������-10</b></option>"
 response.write "<option value='����-50'><b>&nbsp;&nbsp;������-50</b></option>"
 response.write "<option value='����-100'><b>&nbsp;&nbsp;������-100</b></option>"
 response.write "<option value='����ն��'><b>�������ն�ס�</b></option>"
 if lh_set(0)="1" then response.write "<option value='���'><b>����뱻������</b></option>"
 response.write "</select> �Ĵ���</td></tr>"
 response.write "<tr><td colspan=2 align=left class=tablebody1><b>Ͷ�����ݺ�����:</b><br><textarea name=text cols=100 rows=8></textarea><br></td></tr>"
 response.write "<tr><td align=center colspan=3 class=tablebody2>"
 response.write "<input type=submit value='�桡��' name=submit >&nbsp;&nbsp;&nbsp;&nbsp;" 
 response.write "<input type=reset value='�ء�д' name=reset >&nbsp;&nbsp;&nbsp;&nbsp;"
 response.write "<input type=button value='������' onClick='javascript:history.back()' name=button>"
 response.write "<br></td></tr></table></form></CENTER></td></tr></table>"     
end if  
call footer()  
%> 