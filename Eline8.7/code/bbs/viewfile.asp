<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<%
dim Downid
stats="�����ļ�"

if Cint(GroupSetting(49))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û�����������չ����Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
	founderr=true
end if
if request("id")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ָ������ļ���"
	founderr=true
elseif not isInteger(request("id")) then
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ����ļ�������"
	founderr=true
else
	DownID=request("id")
end if
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
	call footer()
else
	set rs=conn.execute("select * from dv_upfile where F_id="&downid)
	if rs.eof and rs.bof then
		call nav()
		call head_var(2,0,"","")
		Errmsg="<br><li>û���ҵ�����ļ���"
		call dvbbs_error()
		call footer()
	else
		if rs("f_flag")=0 then
			conn.execute("update dv_upfile set F_DownNum=F_DownNum+1 where F_ID="&DownID)
			response.redirect "UpLoadFile/"&rs("F_filename")
		end if
	end if
end if
%>