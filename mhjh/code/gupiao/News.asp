<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<meta name=keywords content="魔幻江湖股票市场">
<title><%=Gupiao_Setting(5)%>-<%=membername%></title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>

<table cellspacing=1 cellpadding=0 border=0 width="650" bgcolor="0066CC" align="center">
	<tr> 
		<th height=25 id="TableTitleLink" style="font-weight:normal"><a name="top">股市实时新闻</a><a href="#bottom">[底部]</a><a href="stock.asp" class="cblue">[返回]</a></th>
	</tr>
	<tr> 
		<td width="100%" bgcolor="#444444">
			<table border="0" width="100%">
			<tr><td colspan="4" height="10"></td></tr>
<%
			set rs=server.createobject("adodb.recordset")
			sql="select content,Addtime,id from RndEvent order by ID desc"
			rs.open sql,conn
			do while Not RS.Eof
				response.Write "<tr><td width=5></td><td align=left><font color=white>"&rs(2)&"</font></td><td align=left>"&rs(0)&"</td><td align=right><font color=""#FFCC99"">"&rs(1)&"</font></td><td width=5></td></tr>"
				RS.MoveNext
			Loop
			rs.close
%>				
			<tr><td colspan="4" height="10"></td></tr>	
			</table>
		</td> 
	</tr> 
	<tr>
		<td align=center background="images/title.gif" height=20 bgcolor="#E4E8EF"><a href="stock.asp" class="cblue">[返回股市]</a><a href="#top">[顶部]</a><a name="bottom"></a></td>
	</tr>
</table>
<%CloseDatabase		'关闭数据库%>     