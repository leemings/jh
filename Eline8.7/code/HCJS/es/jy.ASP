<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
if Session("sjjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();}</script>"
	Response.End 
end if

sl=abs(int(request.querystring("sl")))
id=lcase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
name=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from w where id=" & id & " and c<>'卡片' and i>="&sl&" and b="&sjjh_id,conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：赠送不成功，你没有这样的物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
wpname=rs("a")
%>
<title>江湖物品交易♀wWw.51eline.com♀</title>
<body background="../../JHIMG/BK_HC3W.GIF">
<table border=1 bgcolor="#BEE0FC" align=center width=336 cellpadding="10" cellspacing="13" height="293">
<tr valign="top">
<td bgcolor=#cccccc height="226">
<div align="center">
<p><font size="3" color="#000000">物品交易</font><font color="#FF3333" size="3"><br>
</font> </p>
<p><font color="#FF3333" size="3"> </font> </p>
</div>
<table width="324" height="212">
<tr>
<td height="179">
<form method=POST action='jy1.asp?id=<%=rs("id")%>&sl=<%=sl%>'>
<div align="left">
<p><font size="3" color="#000000">物品主</font>人<font color="#0000FF">:<%=name%>
</font><br>
a称<b><font color="#FF0000">:<%=wpname%><br>
</font></b>物品数量<b><font color="#FF0000">:<%=sl%></font></b>
个<br>
<font size="3" color="#000000">交易金额</font>:
<input type=text name=money size=9 maxlength="8" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="100">
<br>
<font size="3" color="#000000">交易姓名</font>:
<input type=text name=name size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="交易人名">
<br>
<font size="3" color="#000000">说明</font>:
<input type="text" name="zy" size="30" value="写上你要说的话" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" maxlength="19">
&nbsp; </p>
<p align="center">
<input type=submit value=确定交易 name="submit" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
<input type=button value=返回上页 onClick="javascript:history.go(-1)" name="button" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
</p>
</div>
</form>
</td>
</tr>
</table>
</td>
</tr>
</table>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>