<%
	dim connshop
	dim strshop
	dim dbshop
	'�����������оݿ�����
	dbshop="../bbs/data/eline_bbs_6.3.0.asp"
	Set connshop = Server.CreateObject("ADODB.Connection")
	strshop="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbshop)
	'�����ķ��������ý��ϰ汾Access�����������������ӷ���
	'strshop="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(dbshop)	
	connshop.Open strshop%>