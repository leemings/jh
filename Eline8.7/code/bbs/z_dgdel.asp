<!--#include file="conn.asp"-->
<!--#include file="z_dgconn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'---------------------
' write by ��ˮ��ɽ
'---------------------
dim id,candeletedg
candeletedg=false	
if membername="" or (not founduser) then
	errmsg=errmsg+"<br>"+"<li>����û�е�¼��̳�����¼������"
	founderr=true
end if	
if request("id")="" or (not isnumeric(request("id"))) then
	errmsg=errmsg+"<br>"+"<li>�����Ƿ�����ȷ�����Ǵ���Ч���ӽ���"
	founderr=true
else
	id=clng(request("id"))
end if
'==========�ж��Ƿ����ɾ����� ��ʼ
if not founderr then
	set rs=connDG.execute("select sender,incept from [media] where id="&id)
	if rs.bof and rs.eof then
		errmsg=errmsg+"<br>"+"<li>�����Ƿ�����ȷ�����Ǵ���Ч���ӽ���"
		founderr=true
	elseif rs(0)=membername or rs(1)=membername or master then
		candeletedg=true
	else
		errmsg=errmsg+"<br>"+"<li>��û��ɾ���˵���Ȩ��"
		founderr=true			
	end if	
end if		
'==========�ж��Ƿ����ɾ����� ����	
if founderr or (not candeletedg) then
	stats="������"
	call nav()
	call head_var(0,0,"���������","z_dglistall.asp")
	call dvbbs_error()
	call CloseDB()
	call footer()
else	
	sql="delete * from [media] where id="&id
	connDG.Execute(sql)
	conn.close
	set conn=nothing
	call CloseDB()
	response.redirect Request.ServerVariables("HTTP_REFERER")
end if
%>