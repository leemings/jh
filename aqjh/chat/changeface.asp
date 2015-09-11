<%@ LANGUAGE=VBScript codepage ="936" %>
<%response.expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
name=aqjh_name
face=left(trim(request.form("face")),3)
if InStr(face,"or")<>0 or InStr(face,"'")<>0 or InStr(face,"`")<>0 or InStr(face,"=")<>0 or InStr(face,"-")<>0 or InStr(face,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
Response.End 
end if
face="../ico/"&face&"-2.gif"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
conn.Execute "update 用户 set 名单头像='"&face&"' where 姓名='" & aqjh_name &"'"
conn.close
set conn=nothing
%><html>
<head>
<title>更改头像</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed" leftmargin="0" topmargin="30">
<p><font color="#FFFFFF">你已经成功的修改了头像<img src="<%=face%>">，退出江湖重新进入您就可以看到你更改后的头像！</font></p>
<p><br>
</p>
</body>
</html>