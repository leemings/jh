<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 银两,金币,性别,保留,hw from 用户 where 姓名='"&aqjh_name&"'",conn
if rs("性别")="男" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('妇科医院，男人止步！');window.close();}</script>"
	Response.End
end if
hai=rs("保留")
yin=rs("银两")
jb=rs("金币")
rs.close
if hai="保留" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你还没有身孕呢，保个什么胎呀！');window.close();}</script>"
	Response.End
end if
hdata=split(hai,"|")
baby=hdata(0)
babysj=hdata(1)
babytl=clng(hdata(2))
babyzt=hdata(3)
erase hdata
if babyzt="正常" or babyzt="药1" or babyzt="针1" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的胎儿一切正常，暂时不需要进行任何保育措施！');window.close();}</script>"
	Response.End
end if
if babyzt="生" or babyzt="难" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('哎呀，你就要生了，产房在隔壁，快到产房去呀！');window.close();}</script>"
	Response.End
end if
if babyzt="药" then
	yin1=5000000
	jb1=15
	if yin<yin1 or jb<jb1 then
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你需要使用保胎药，费用：银两500万，金币15个，你的钱不够！');window.close();}</script>"
		Response.End
	end if
	babyzt="药1"
	babytl=babytl+300
	h="<img src=images/yao.gif width=100 height=165><br><br>医生给你服用了保胎药，胎儿生命活力增加300点。收费500万及15金币。"
end if
if babyzt="针" then
	yin1=15000000
	jb1=20
	if yin<yin1 or jb<jb1 then
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你需要打保胎针，费用：银两1500万，金币20个，你的钱不够！');window.close();}</script>"
		Response.End
	end if
	babyzt="针1"
	babytl=babytl+900
	h="<img src=images/z.gif width=100 height=165><br><br>医生为你打了一支保胎针，胎儿生命活力增加900点，收费银两1500万及20金币。"
end if
newbl=baby&"|"&babysj&"|"&babytl&"|"&babyzt
conn.execute "update 用户 set 银两=银两-"&yin1&",金币=金币-"&jb1&",保留='"&newbl&"' where 姓名='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>保育科</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
 
<div align="center"><br>
  <br>
  <font size="6" color="#FF0000">妇科中心_保育科</font><br>
  <br>
  <br>
  <%=h%></div>
</body>
</html>