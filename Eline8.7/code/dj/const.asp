<%Application.Lock
set rs=server.createobject("adodb.recordset")
sql="select * from const" 
rs.open sql,conn,1,1
if not rs.eof then
	Application("WebSiteName")=rs("WebSiteName")		    	'��վ����
	Application("WebSiteUrl")=rs("WebSiteUrl")		        	'��վ��ַ
	Application("Logo")=rs("Logo")					        	'��վLOGO
	Application("WebMaster")=rs("WebMaster")					'վ������
	Application("oicq")=rs("oicq")								'վ��QQ
	Application("EMail")=rs("EMail")				 			'վ������
	Application("newreg")=rs("newreg")		                 	'�Ƿ������»�Աע��
	Application("collect")=rs("collect")
	Application("FreeListen")=rs("FreeListen")
	Application("Urgent")=rs("Urgent")
	Application("djserver")=rs("djserver")
end if
rs.close
set rs=nothing
Application.UnLock
	WebSiteName=Application("WebSiteName")				'��վ����
	WebSiteUrl=Application("WebSiteUrl")				'��վ��ַ
	Logo=Application("logo")							'��վLOGO
    Oicq=Application("Oicq")							'վ��OICQ����
	EMail=Application("EMail")							'վ������
	newreg=Application("newreg")		             	'�Ƿ������»�Աע��
	WebMaster=Application("WebMaster")					'վ������
	collect=Application("collect")
	FreeListen=Application("FreeListen")
	Urgent=Application("Urgent")
	djserver=Application("djserver")
%>