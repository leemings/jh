<!--#include file="conn.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_const.asp"-->
<%
response.buffer=true
stats="��Ժ��������"
call nav()
call head_var(0,0,"������Ժ","z_fy_fayuan.asp")
if not master then 
  Errmsg=Errmsg+"<br>"+"<li>��û�н��뱾������Ժ��̨��Ȩ�ޣ���ʹ���˷Ƿ�������"
   call dvbbs_error()
else
   call getfyconfig()
   if request("action")="save" then
     bs_set=request.form("bsbz")+"|"+request.form("bsmoney")
     tj_set=request.form("tjbz")+"|"+request.form("tjmoney")
     pst_set=request.form("pstbz")+"|"+request.form("pstID")
     fj_set=request.form("fj1")+","+ request.form("fj2")+","+ request.form("fj3")+","+ request.form("fj4")+","+request.form("fj5") 
     lh_set=request.form("lhbz")+"|"+request.form("lhmlbz")+"|"+request.form("lhml")
     qj_set=request.form("qjbz")+"|"+ request.form("qjmlbz")+"|"+request.form("qjml")+"|"+request.form("qjxm")+"|"+request.form("qjtime")
     log_set=request.form("logbz")+"|"+request.form("lognum")
     if isnull(request.form("fgname")) or request.form("fgname")="" then
       fyname="��"
     else
       fyname=request.form("fgname")
     end if
     if isnull(request.form("jyzname")) or request.form("jyzname")="" then
       jyname="��"
     else
       jyname=request.form("jyzname")
     end if
     sql="update [fyconfig] set fyreadme='"&request.form("text")&"',bsmoney='"&bs_set&"',tjmoney='"&tj_set&"',pst='"&pst_set&"',fj='"&fj_set&"',lh='"&lh_set&"',qj='"&qj_set&"',fyadmin='"&fyname&"',jyadmin='"&jyname&"',log='"&log_set&"'" 
     connfy.execute sql
     Response.Redirect "z_fy_fayuan.asp"
  else
   response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=25><b>��Ժ������������</b></td></tr><tr><td valign=middle class=tablebody2 height=100><CENTER>"
   response.write "<form action=z_fy_manage.asp?action=save method=post>"
   response.write "<table class=tableborder1 cellspacing=1 cellpadding=3 align=center ><tr>"
   response.write "<th align=left colspan=2>��Ժ��������</th><th align=right colspan=2>��������:<input name=fgname size=18 maxlength=20 value="
   if isnull(fyname) or fyname="" then 
      response.write "��"
   else
      response.write fyname
   end if
   response.write "></th></tr>"   
   response.write "<tr><td class=tablebody2 width=15% >��Ժ����˵����</td><td colspan=3 class=tablebody1><textarea name=text cols=100 rows=8 >"&rs(0)&"</textarea></td>"
   response.write "<tr><td class=tablebody2>���ӽ����趨��<br>��ʽΪ <b>ԭ��|����</b><br>����ǰ�ӣ��ţ�<br>����ǰ�ӣ��ţ�</td><td colspan=3 class=tablebody1>ͬ�⣺<input name=fj1 size=8 value="&fj_set(0)&"> ���أ�<input name=fj2 size=8 value="&fj_set(1)&"> ƽ����<input name=fj3 size=8 value="&fj_set(2)&"> �ܸ棺<input name=fj4 size=8 value="&fj_set(3)&"> �����ܸ棺<input name=fj5 size=8 value="&fj_set(4)&"></td></tr>"
   response.write "<tr><td class=tablebody2>���������Ź��ܣ�</td><td class=tablebody1><input type=radio name=pstbz value=1 "
   if pst_set(0)="1" then response.write " checked"
   response.write ">����<input type=radio name=pstbz value=0 "
   if pst_set(0)="0" then response.write " checked"
   response.write ">�ر�</td>"
   response.write "<td colspan=2 class=tablebody1>�������BoardID��&nbsp;<input name=pstID value="&pst_set(1)&" size=18> �趨�ύ������ͶƱ���ӷ������ID</td></tr>"
   response.write "<tr><td class=tablebody2>������鹦�ܣ�</td><td class=tablebody1><input type=radio name=lhbz value=1 "
   if lh_set(0)="1" then response.write " checked"
   response.write ">����<input type=radio name=lhbz value=0 "
   if lh_set(0)="0" then response.write " checked"
   response.write ">�ر�</td>"
   response.write "<td width=15% class=tablebody1>����Ƿ��������</td>"
   response.write "<td class=tablebody1><input type=radio name=lhmlbz value=1 "
   if lh_set(1)="1" then response.write " checked"
   response.write ">��<input type=radio name=lhmlbz value=0 "
   if lh_set(1)="0" then response.write " checked"
   response.write ">���ۡ�  �۳�����ֵ<input name=lhml size=18 maxlength=20 value="&lh_set(2)&"> �Ƿ���������</td></tr>"   
   response.write "<th align=left colspan=2>������������</th><th align=right colspan=2>����������:<input name=jyzname size=18 maxlength=20 value="
   if isnull(jyname) or jyname="" then 
      response.write "��"
   else
      response.write jyname
   end if
   response.write "></th></tr>"   
   response.write "<tr><td class=tablebody2>���ű��͹��ܣ�</td><td class=tablebody1><input type=radio name=bsbz value=1" 
   if bs_set(0)="1" then response.write " checked"
   response.write ">����<input type=radio name=bsbz value=0"
   if bs_set(0)="0" then response.write " checked"
   response.write " >�ر�</td>"
   response.write "<td class=tablebody1 colspan=2>��ͱ��ͽ�<input name=bsmoney size=18 maxlength=20 value="&bs_set(1)&"> �趨��Ѻ��������������ͱ��ͽ����9��</td></tr>"
   response.write "<tr><td class=tablebody2>����̽�๦�ܣ�</td><td width=15% class=tablebody1><input type=radio name=tjbz value=1" 
   if tj_set(0)="1" then response.write " checked"
   response.write ">����<input type=radio name=tjbz value=0"
   if tj_set(0)="0" then response.write " checked"
   response.write " >�ر�</td>"
   response.write "<td class=tablebody1 colspan=2>̽�����Էѣ�<input name=tjmoney size=18 maxlength=20 value="&tj_set(1)&"> �趨̽����Ѻ��������������ȡ�ֽ����9ǧ��</td></tr>"   
   response.write "<tr><td class=tablebody2>�Զ������¼���</td><td width=15% class=tablebody1><input type=radio name=logbz value=1" 
   if log_set(0)="1" then response.write " checked"
   response.write ">����<input type=radio name=logbz value=0"
   if log_set(0)="0" then response.write " checked"
   response.write " >�ر�</td>"
   response.write "<td class=tablebody1 colspan=2>�¼���������<input name=lognum size=18 maxlength=20 value="&log_set(1)&"> �趨���������ٺ�̽���¼��Զ�������</td></tr>"   
   response.write "<tr><td class=tablebody2 rowspan=2>���ٹ������ã�</td>"
   response.write "<td width=15% class=tablebody1>����������������</td>"
   response.write "<td class=tablebody1><input type=radio name=qjbz value=1 "
   if qj_set(0)="1" then response.write " checked"
   response.write ">����<input type=radio name=qjbz value=0 "
   if qj_set(0)="0" then response.write " checked"
   response.write ">�ر�</td>"
   response.write "<td class=tablebody1>�����������Ƿ��������ֵ��&nbsp;&nbsp;<input type=radio name=qjmlbz value=1 "
   if qj_set(1)="1" then response.write " checked"
   response.write ">�۳�<input type=radio name=qjmlbz value=0 "
   if qj_set(1)="0" then response.write " checked"
   response.write ">���ۡ�  ����ֵ<input name=qjml size=4 maxlength=4 value="&qj_set(2)&"> </td></tr>"
   response.write "<tr><td class=tablebody1>��ʾ������������</td>"
   response.write "<td class=tablebody1><input type=radio name=qjxm value=1 "
   if qj_set(3)="1" then response.write " checked"
   response.write ">��ʾ<input type=radio name=qjxm value=0 "
   if qj_set(3)="0" then response.write " checked"
   response.write ">�ر�</td>"
   response.write "<td class=tablebody1>�Ƿ��޶�20�������ٴ������ƣ�<input type=radio name=qjtime value=1 "
   if qj_set(4)="1" then response.write " checked"
   response.write ">����<input type=radio name=qjtime value=0 "
   if qj_set(4)="0" then response.write " checked"
   response.write ">����"
   response.write "</td></tr>"
response.write "<tr><td class=tablebody2 colspan=4><div align=center><input type=submit name=Submit value='�����޸�'></div></td></tr>"
   response.write "</table></form></CENTER></td></tr></table>"
  end if 
end if
call footer() 
%>
