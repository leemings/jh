<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
pid=Request.QueryString("id")
if pid="" or not isnumeric(pid) then Response.Redirect "../error.asp?id=024"
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from pet where username='"&username&"' and exist=true and id="&pid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=069"
sinew=rst("sinew")
maxsinew=rst("maxsinew")
cleanliness=rst("cleanliness")
jiankang=rst("health")
kuaile=rst("happy")
rst.Close
set rst=nothing
grow=sinew*10\maxsinew+cleanliness+jiankang+kuaile
if grow<320 then Response.Redirect "../error.asp?id=071"
conn.Execute "update pet set sinew="&maxsinew&",health=100,happy=100,cleanliness=100,exist=false,option_M='放生',option_T="&nowtimetype&" where id="&pid
conn.Execute "update 用户 set 道德=道德+"&grow&",银两=银两+"&maxsinew*3&" where 姓名='"&username&"'"
conn.Close
set conn=nothing
%>
<html>
<head>
<title>宠物之家</title>
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false '>
<p align=center><font color="0000ff" size="3">宠物之家</font><br>这是心的呼唤,这是爱的交流</p><hr>
<p align=center>
放生成功.您的银两加<%=maxsinew%>,道德加<%=grow%>:)<br>
<input type=button value='返回' onclick=javascript:location.href='pet.asp';>
</p>
</body>
</html>
