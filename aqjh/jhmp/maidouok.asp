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
dds=request.form("dds")
if not isnumeric(dds) then 
	Response.Write "<script language=JavaScript>{alert('提示：["&dds&"]输入错误，豆子数量请使用数字！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if instr(dds,"'")<>0 or instr(dds," ")<>0 or instr(dds,"or")<>0 or instr(dds,"OR")<>0 or instr(dds,"%20")<>0 or instr(dds,"=")<>0 or instr(dds,">")<>0 or instr(dds,"<")<>0 then Response.Redirect "../error.asp?id=440"
dds=int(abs(dds))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 泡豆点数 from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
end if
pdds=rs("泡豆点数")
rs.close
if pdds<1000 then
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('你现在还没有豆子，只有："&pdds&"点，不能卖豆子！');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
dousl=int(pdds/1000)
if dds>dousl then
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('你有这么多豆子吗？！');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
d=dds*1000
s=pdds-dds*1000
conn.execute "update 用户 set 泡豆点数="& s &",会员金卡=会员金卡+"& 2*dds &" where 姓名='"&aqjh_name&"'"
mess="你卖掉豆子："&dds&"个，会员费增加："&dds*2&"元，还剩:"&s&"点！"
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>卖豆子</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="../bg.gif">
<div align="center"><font color="#FF0000"><%=mess%></font> 
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p><input type=button value=关闭窗口 onClick='window.close()' name="button"></p>
</div>
</body>
</html>