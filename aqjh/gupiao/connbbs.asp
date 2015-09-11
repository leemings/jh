<%
dim connbbs
	dim connstrbbs
	dim dbbbs
	dim sqlbbs,rsbbs
	 
	dbbbs=Application("aqjh_usermdb")
	Set connbbs = Server.CreateObject("ADODB.Connection")
	connstrbbs="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbbbs)
 	connbbs.Open connstrbbs
%>