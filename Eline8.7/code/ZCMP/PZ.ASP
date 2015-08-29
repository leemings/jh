<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then
	Response.Write "<script Language=Javascript>alert('你不是站长？！');window.close();</script>"
	response.end
end if
id=trim(request("id"))
if left(id,3)="%20" OR InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then
	Response.Redirect "../error.asp?id=120"
	response.end
end if
%><!--#include file="data.asp"--><%
sql="select 申请人,姓名 from s where 申请人="&id
set rs=conn1.execute(sql)
if rs.eof or rs.bof then
	set rs=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('此人还未申请过！');window.close();</script>"
	response.end
end if
you=rs("姓名")
rs.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 等级,魅力,道德,银两,门派 from 用户 where id="&id,conn
if rs.eof or rs.bof then
	set rs=nothing
	conn.close
	set conn=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('江湖中找不到这个人！');window.close();</script>"
	response.end
end if
you=rs("姓名")
tmprs=conn.execute("Select count(*) As 数量 from 用户 where 等级>=10 and 介绍人='"& you &"'")
lr=tmprs("数量")
set tmprs=nothing
if rs("等级")<65 or rs("魅力")<100000 or rs("道德")<100000 or rs("银两")<1000000000 or lr<20 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('这个人的各项指数还不够，他是怎么申请上的？请核实后再说吧！');window.close();</script>"
	response.end
end if
rs.close
set rs=nothing
conn.execute "update 用户 set grade=6,门派='官府',身份='弟子' where id="&id
sql="update s set 申请结果='已批准',审批时间=now() where 申请人="&id
set rs1=conn1.execute(sql)
conn.close
set conn=nothing
conn1.close
set conn1=nothing
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>你成功把</title>
</head>

<body background="../JHIMG/BK_HC3W01.GIF">

<p align="center"><b><font size="5">你成功把<%=you%>加为6级官府</font></b></p>

<p align="center"><b><font size="5"><a href="gfsq.asp">返回</a></font></b></p>

</body>

</html>
