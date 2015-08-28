<%
dim connbbs
	dim connstrbbs
	dim dbbbs
	dim sqlbbs,rsbbs
	 
	dbbbs=Application("yx8_mhjh_connstr")
	Set connbbs = Server.CreateObject("ADODB.Connection")
	connstrbbs="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbbbs)
 	connbbs.Open connstrbbs
%>