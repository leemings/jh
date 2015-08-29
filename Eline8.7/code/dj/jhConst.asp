<%	
dim membername
	membername=session("sjjh_name")
	
	if membername="" then
		response.write "<title>一线视听-提示信息</title><style type=text/css>td {  font-family: 宋体; font-size: 9pt}</style>"
		errmess="<li>您还没有登陆，请登陆后再来！"
		call endinfo(0)
		response.end
	end if
%>