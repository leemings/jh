<%
dim jhname
dim jhdj
jhname=session("sjjh_name")
jhdj=Session("sjjh_jhdj")
if jhname="" then
response.write "<script Language=Javascript>alert('�Բ�������û�е�½�����½��������');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
%>