<%
function bangfu(un,st,mg)
rst.Open "select ����,����¼IP from �û� where ����='"&st&"'",conn
  dj=rst("����")
  dl=rst("����¼IP")
  rst.close
  if dj>500 then
	Response.Write "<script language=JavaScript>{alert('�Է����������ˣ��㲻���پȼ����ˣ�');}</script>"
	Response.End
	exit function
  end if
if not isnumeric(mg) then
bangfu="<font color=FF0000>��������ˡ�</font>##�����ȷ����Ǯ"&mg&"��%%��"
else
mg=clng(mg)
if mg<1000000 then
bangfu="<font color=FF0000>��������ˡ�</font>##�����ȷ����Ǯ"&mg&"��%%������100�򣬱�̫���ģ�"
elseif mg>3000000 then
bangfu="<font color=FF0000>��������ˡ�</font>##���ڹٸ���ֹ��Ǯ�����˳���300��лл��ĺ��ģ�"
else
rst.Open "select ����,����ʱ��,����¼IP from �û� where ����='"&un&"' and ����>="&mg,conn
if rst.EOF or rst.BOF then
bangfu="<font color=FF0000>��������ˡ�</font>%%Ц�Ŷ�##˵��'�����������⣬��������"&mg&"�������������Һ���'"
else
    sj=rst("����¼IP")
    rst.close
if sj=dl then
	Response.Write "<script language=JavaScript>{alert('��ͬ���������IP�޷����а�����˻��');}</script>"
	Response.End
exit function
	end if
conn.execute "update �û� set ����=����+"&mg*0.97&",����=����+500,����=����+"&mg/2&",����=����+"&mg/5&" where ����='"&st&"'"
conn.execute "update �û� set ����=����-"&mg&",����=����-"&mg/2&",����=����-"&mg/5&",����=����+"&mg/50&",����=����+"&mg/50&" where ����='"&un&"'"
bangfu="<font color=FF0000>��������ˡ�</font>##����ħ�ý�����������ͳ�����İ������ˣ����Լ���"&mg&"�����ӡ�"&mg/2&"������"&mg/5&"�����͸���%%�������Լ���ùٸ���������"&mg/50&"������"&mg/50&"����ٸ��������Ľ���˰��3%<bgsound src='../mid/thanks.wav' loop=1>"
end if
end if
end if
end function
%>