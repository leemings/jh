<!-- #include file="../const5.asp" -->
<%
dim jhname
dim jhdj
jhname=session("aqjh_name")
jhdj=Session("aqjh_jhdj")
if jhname="" or jhdj<bbs_add_dj then
response.write "<script Language=Javascript>alert('���棺�ȼ�����������û�е�½��');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(jhname)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('���棺���ȵ�½�����ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>