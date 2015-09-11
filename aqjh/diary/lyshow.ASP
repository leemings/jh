<%@ LANGUAGE=VBScript codepage ="936" %>
<%aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
id=request("id")
if id="" or (not isnumeric(id)) then 
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你在搞什么！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
rs1.open "select * from 留言 where id="&id,conn1,1,1
aa1=rs1("姓名")
aa2=badstr(rs1("标题"))
aa3=badstr(rs1("内容"))
aa4=rs1("时间")
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing
function badstr(str)
	badword="射精,奸,屎,妈,娘,尻,操,王八,逼,贱,狗,婊,叉你,插你,干你,鸡巴,睾丸,蛋,包皮,龟头,屄,赑,妣,肏,奶,尻,屌,作爱,做爱,床,抱抱,鸡八,处女,打炮,十八摸,爸,我儿,·,主席,泽民,法伦,洪志,大法,公安,政府,反动,法院,升天,周天,功"
	bad=split(badword,",")
	for i=0 to ubound(bad)-1
		if InStr(LCase(str),bad(i))<>0 then 
			str=replace(str,bad(i),"*")
		end if
	next
	badstr=str
end function
%>
<html><head><title>查看留言</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head><LINK href="images/diary.css" rel=stylesheet type=text/css>
<body bgcolor="#88afd7" text="#FFFF00">
<div align="center"><font color="#FFFF00" size="+2" face="楷体_GB2312"><b>查看留言</b></font><br>
</div>
<table width="100%" border="1" cellspacing="0" cellpadding="3">
  <tr> 
    <td width="19%">标题：</td>
    <td><%=aa2%></td>
  </tr>
  <tr>
    <td width="19%" height="9">内容：</td>
    <td height="9"><%=aa3%></td>
  </tr>
  <tr>
    <td width="19%">作者：</td>
    <td><%=aa1%> 时间:<%=aa4%></td>
  </tr>
</table>
<div align="center"><br>
  <input type="button" value="关闭窗口" onClick="window.close()"
name="button">
</div>
