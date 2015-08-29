<%
dim z_photo_cn,z_photo_connstr,z_photo_db

'更改数据库名字
z_photo_db="data/e_photo.asp"
Set z_photo_cn = Server.CreateObject("ADODB.Connection")
z_photo_connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(z_photo_db)
z_photo_cn.Open z_photo_connstr

function Close_z_photo_cn
	z_photo_cn.close
	Set z_photo_cn = Nothing
End Function
%>
