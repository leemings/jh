<%	
dim membername
	membername=session("sjjh_name")
	
	if membername="" then
		response.write "<title>E�߽�����վ-FlashƵ��-��ʾ��Ϣ</title><style type=text/css>td {  font-family: ����; font-size: 9pt}</style>"
		errmess="<li>����û�е�½�����½��������"
		call endinfo(0)
		response.end
	end if
%>