<!--#include file="conn.asp"-->
<!--#include file="Z_TestConn.ASP"-->
<!-- #include file="inc/const.asp" -->
<%
	dim ars,auserjl,azql,adts,afjjl,auserupsinew,auserarticle
	dim AnswerSetting,OtherSetting
	stats="���Ĵǵ� ֪ʶ�ʴ����"
	call nav()
	call head_var(0,0,"���Ĵǵ�","Z_test.asp")
	call activeonline()
	
	sql="select AnswerSetting,OtherSetting,userjl,zql,dts,fjjl,userarticle,userupsinew from [config]"
	set ars=aconn.execute(sql)
	AnswerSetting=split(ars(0),"||")	'0ÿ�ش�һ��������Ϸ�Ҹ��� ||1ÿ�ش�һ������ֵ����ֵ||2ÿ�ش�һ�⾭��ֵ����ֵ||3������֮���ʱ����||4������||5�Ƿ���ʾ��ȷ��||6��Ǯ��֪ʶ�����Ǯ||7�����Ƿ���ʱ||8��Ŀ�Ƿ�֧��UBB
	if ubound(AnswerSetting)<8 then
		redim preserve AnswerSetting(8)
		AnswerSetting(8)="1"
	end if
	OtherSetting=split(ars(1),"||")		'0�Ƿ���ʾͼƬ || 1ͼƬ·�� || 2������������ʾ��¼����
	auserjl=ars(2)			'�ϴ�������� 
	azql=ars(3)				'��ø��ӽ������� ���δ�����ȷ��
	adts=ars(4)				'��ø��ӽ������� �������ٴ�����
	afjjl=ars(5)			'���ӽ������
	auserarticle=ars(6)		'ÿ��һ������Ϸ��
	auserupsinew=ars(7)		'�ϴ�һ�⽱��Ϸ��   
	ars.close
	set ars=nothing
%>
<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
<tr><th>���Ĵǵ�---֪ʶ�ʴ����</th></tr>
<tr><td class=tablebody2>

  <table align=center cellspacing=1 cellpadding="3" class=tableborder1 style="width:80%">
    <tr>
          <td width="100%" class=TopLighNav1>
          <p align="center"><b><font color="#FF0000">֪ʶ�ʴ����</font></b></td>
    </tr>
    <tr>
          <td width="100%" class=tablebody2>
		  <br>
		  <p align="center"><font color=blue>���Ĵǵ� V1.2 ��</font></p>
		  &nbsp;<span lang="zh-cn">1�����������Ϸ�ң�</span><p>
          <span lang="zh-cn">����A����������Ϸ�ң�ÿ��һ����������Ϸ�� <font color=red><b><%=auserarticle%></b></font> ö</span>,<span lang="zh-cn">�Զ����¡�</span></p>
          <p><span lang="zh-cn">����B���ϴ���Ŀ����Ϸ�ң�ÿ�ϴ�һ����Ŀ��ͨ����ˣ�������Ϸ�� <font color=red><b><%=auserupsinew%></b></font> ö��</span></p>
          <p>&nbsp;<span lang="zh-cn">2��֪ʶ�ʴ��й����ò�����</span></p>
          <p><span lang="zh-cn">������Ϸ�ң�&nbsp;ÿ��һ��<b>&nbsp;<font color="#FF0000">��<%=AnswerSetting(0)%></font></b></span></p> 
          <p><span lang="zh-cn">��������ֵ��</span><span lang="en-us">&nbsp;</span>��<span lang="en-us"><font color="#FF0000"><b> +</b></font></span><b><font color="#FF0000"><%=AnswerSetting(2)%></font></b> ��<b><font color="#FF0000"><span lang="en-us"> -</span><%=AnswerSetting(2)%></font></b></p>
          <p><span lang="zh-cn">��������ֵ��</span><span lang="en-us">&nbsp;</span>��<b><font color="#FF0000"><span lang="en-us"> +</span><%=AnswerSetting(1)%></font></b> ��<b><font color="#FF0000"><span lang="en-us"> -</span><%=AnswerSetting(1)%></font></b></p>
          <p><span lang="zh-cn">�����֡���</span><font color=red>����Ŀ��ͬ��Ӧ����</font></p>
          <p><span lang="zh-cn">����ÿ�δ�������</span><font color=red><b><%=AnswerSetting(4)%></b></font></p>
          <p><span lang="zh-cn">����ÿ�δ���ʱ������</span><font color=red><b><%=AnswerSetting(3)%></b></font></p>
          <p><span lang="zh-cn">�����õ����ӽ���<font color=red><b><%=afjjl%></b></font>Ԫ������������������ڣ�<font color=red><b><%=adts%></b></font>  ��ȷ�ʲ����ڣ�</span><font color=red><b><%=azql%></b>%</font></p>
          <p><span lang="zh-cn">������ʾ��ȷ�𰸣�</span><font color=red><%if cint(AnswerSetting(5))=1 then%>��ʾ<%else%>��Ǯ��֪ʶ<%end if%></font></p> 
		  <p><span lang="zh-cn">�������������Ǯ��</span><font color=red><%=AnswerSetting(6)%> Ԫ</font></p>
		  <p><span lang="zh-cn">��������ʱ�����ƣ�</span><font color=red><%if cint(AnswerSetting(7))=1 then%>��<%else%>�ر�<%end if%></font></p>
		  <p><span lang="zh-cn">������Ŀ�Ƿ�֧��UBB��ǩ��</span><font color=red><%if cint(AnswerSetting(8))=1 then%>֧��<%else%>��֧��<%end if%></font></p>
		  <p><span lang="zh-cn">��������˳��</span><font color=red>���˳��</font></p>
		  &nbsp;<span lang="zh-cn">3����ο�ʼ��Ϸ��</span><p> 
		  <p><span lang="zh-cn">���������Ŀ���ͼ��ɽ���֪ʶ�ʴ����</span></p> 
		  <p><span lang="zh-cn">����</span></p>     
		  </td>
    </tr> 
    <tr>
    <td width="100%" align="center" class=tabletitle2><a href=javascript:history.go(-1)>������һҳ</a></td> 
    </tr>
</table>
</td></tr></table>
<%
set aconn=nothing
call footer()
%>