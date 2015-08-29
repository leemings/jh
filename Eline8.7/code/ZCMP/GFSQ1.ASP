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

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 等级,魅力,道德,银两,门派 from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=440"
	response.end
end if
sjjh_id=rs("id")
if rs("等级")<65 or rs("魅力")<100000 or rs("道德")<100000 or rs("银两")<1000000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=460"
	response.end
end if
if rs("门派")="官府" then
	Response.Write "<script Language=Javascript>alert('提示：你现在已经是官府了！');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
tmprs=conn.execute("Select count(*) As 数量 from 用户 where 等级>=20 and 介绍人='"& sjjh_name &"'")
lr=tmprs("数量")
set tmprs=nothing
if lr<20 then
	Response.Write "<script Language=Javascript>alert('提示：你的拉人记录不足20人，或你所拉的人的等级还没有到20级以上！');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 会员等级 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,3,3
hy=rs("会员等级")
if hy=rs("会员等级")and (hy=1 or hy=2 or hy=3 or hy=4 or hy=5) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('等级会员无权自创门派!');window.close();}</script>"
		response.end
end if
rs.close
'e加入限制
%><!--#include file="data.asp"--><%
sql="select 申请人 from s where 申请人="&sjjh_id
set rs1=conn1.execute(sql)
if rs1.eof or rs1.bof then
	sql="insert into s(申请人,姓名,申请时间,审批时间) values ('"&sjjh_id&"','"&sjjh_name&"',now(),now())"
	set rs1=conn1.execute(sql)
	set rs1=nothing
	conn.close
	set conn=nothing
	conn1.close
	set conn1=nothing
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body {font-size:9pt;SCROLLBAR-FACE-COLOR: #c2cea8;SCROLLBAR-HIGHLIGHT-COLOR: #ffffff;SCROLLBAR-SHADOW-COLOR: #000000;SCROLLBAR-3DLIGHT-COLOR: #ffffff;SCROLLBAR-ARROW-COLOR: #ffffff;SCROLLBAR-TRACK-COLOR: #ffffff;SCROLLBAR-DARKSHADOW-COLOR: #ffffff;SCROLLBAR-BASE-COLOR: #c2cea8}
table {font-size: 9pt}
a:link {CURSOR:crosshair;color: yellow; text-decoration: none}
a:visited {CURSOR:crosshair;color: yellow; text-decoration: none}
a:active {CURSOR:crosshair;color: #white; text-decoration: none}
a:hover { CURSOR:crosshair;color: #white; text-decoration: underline overline}
-->
</style>
<script language="javascript">
function opench(url,win){
        controlWindow=window.open(url,win,"fullscreen=no,toolbar=yes,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=20,left=2,width=450,height=300");
        controlWindow.focus();
        }
function ch(){
	document.forms[0].userpassword.value=document.forms[0].pwd.value;
	document.forms[0].pwd.value="";
//	document.forms[0].submit();
	}
</script>

</head>

<body oncontextmenu=self.event.returnValue=false bgcolor=#339966 text=#ffffff link="#000080" alink="#800000" vlink="#000080">
<div align="center"> 
  <p><font size="3"><b>你的申请完成，请等待站长批准！</b></font></p>
</div>
</body>
</html>
<%
else
	rs1.close
	set rs1=nothing
	set rs=nothing
	conn.close
	set conn=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('提示：你已经申请过官府了，等着站长审批吧！');window.close();</script>"
	response.end
end if
%>