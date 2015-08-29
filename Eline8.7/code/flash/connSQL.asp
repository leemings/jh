<%
dim conn,connstr,username,password,database,path
username="eline"			'数据库登陆名
password="e_flash"			'数据库密码
database="eline_flash"			'数据库名
path="(local)"			'数据库所在的IP地址，如果是本地数据库则不用更改
set conn=Server.CreateObject("adodb.Connection")
connstr="PROVIDER=SQLOLEDB;DATA SOURCE="&path&";UID="&username&";PWD="&password&";DATABASE="&database
conn.open connstr
%>