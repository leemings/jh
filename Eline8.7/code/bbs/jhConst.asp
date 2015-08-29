<%
dim jhname
dim jhdj
jhname=session("sjjh_name")
jhdj=Session("sjjh_jhdj")
if jhname="" then
response.write "<script Language=Javascript>alert('对不起，您还没有登陆，请登陆后再来！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
%>