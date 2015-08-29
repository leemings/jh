<%
	dim connDG,dgconnstr
	
	'更改点歌数据库名字
	const dgDB="data/e_dgData.asp" 
	dgconnstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dgDB) 
	
	Set connDG = Server.CreateObject("ADODB.Connection")             	
	connDG.Open dgconnstr
	set dgconnstr=nothing

function CloseDB()
	connDG.close
	set connDG=nothing
End Function	 
%>