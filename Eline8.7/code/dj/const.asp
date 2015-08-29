<%Application.Lock
set rs=server.createobject("adodb.recordset")
sql="select * from const" 
rs.open sql,conn,1,1
if not rs.eof then
	Application("WebSiteName")=rs("WebSiteName")		    	'网站名称
	Application("WebSiteUrl")=rs("WebSiteUrl")		        	'网站地址
	Application("Logo")=rs("Logo")					        	'网站LOGO
	Application("WebMaster")=rs("WebMaster")					'站长姓名
	Application("oicq")=rs("oicq")								'站长QQ
	Application("EMail")=rs("EMail")				 			'站长信箱
	Application("newreg")=rs("newreg")		                 	'是否允许新会员注册
	Application("collect")=rs("collect")
	Application("FreeListen")=rs("FreeListen")
	Application("Urgent")=rs("Urgent")
	Application("djserver")=rs("djserver")
end if
rs.close
set rs=nothing
Application.UnLock
	WebSiteName=Application("WebSiteName")				'网站名称
	WebSiteUrl=Application("WebSiteUrl")				'网站地址
	Logo=Application("logo")							'网站LOGO
    Oicq=Application("Oicq")							'站长OICQ号码
	EMail=Application("EMail")							'站长信箱
	newreg=Application("newreg")		             	'是否允许新会员注册
	WebMaster=Application("WebMaster")					'站长姓名
	collect=Application("collect")
	FreeListen=Application("FreeListen")
	Urgent=Application("Urgent")
	djserver=Application("djserver")
%>