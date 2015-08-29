<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<%
dim yjsk_rst,yjsk_sql
dim codeboxtext,codeboxlogo
set yjsk_rst=Server.CreateObject("ADODB.RecordSet")
yjsk_sql="select * from bbslink where isnull(url) or url=''"
yjsk_rst.Open yjsk_sql,conn,1,1
if yjsk_rst.EOF OR yjsk_rst.BOF then
	errmsg=errmsg+"<br>"+"<li>管理员还没有添加本站的信息，请耐心等待！"
	founderr=true
else
	codeboxtext="&lt;a href="&forum_info(1)&" title="&yjsk_rst("readme")&" target=_blank&gt;"&forum_info(0)&"&lt;/a&gt;"
	if yjsk_rst("logo")="" then
	  codeboxlogo="本站LOGO正在制作中，请选择文字链接吧！"
	else
	  codeboxlogo="&lt;a href="&forum_info(1)&" title="&yjsk_rst("readme")&" target=_blank&gt;"&"&lt;img src="&yjsk_rst("logo")&" border=0 height=31 width=88&gt;&lt;/a&gt;"
	end if
end if
stats="查看连接信息"
call nav()
call head_var(2,0,"","")
if founderr=true then
	call dvbbs_error()
else
	%>
	<table cellpadding=3 cellspacing=1 class=tableborder1 align="center">
		<tr align="center"> 
			<th colspan="2" class=tablebody1>站点信息</th>      
		</tr>
		<tr> 
			<td width="40%" class=tablebody1>论坛名称：</td>
			<td class=tablebody1><%=forum_info(0)%></td>
		</tr>
		<tr> 
			<td width="40%" class=tablebody1>论坛LOGO：</td>
			<td class=tablebody1><%
				if yjsk_rst("logo")="" then
					%>本站LOGO正在制作中……<%
				else
					%><img border="0" src="<%=yjsk_rst("logo")%>" width="88" height="31"><%
				end if
			%></td>
		</tr>
		<tr> 
			<td width="40%" class=tablebody1>连接 URL：</td>      
			<td class=tablebody1><%=forum_info(1)%></td>
		</tr>
	  <tr> 
			<td width="40%" class=tablebody1>连接LOGO：</td>
			<td class=tablebody1><%
				if yjsk_rst("logo")="" then
					%>本站的LOGO正在制作中……<%
				else 
					%><%=yjsk_rst("logo")%><%
				end if
			%></td>
		</tr>
		<tr> 
			<td width="40%" class=tablebody1>论坛简介：</td>
			<td class=tablebody1><%=yjsk_rst("readme")%></td>
		</tr>
		<tr> 
			<td width="40%" class=tablebody1>获得代码：</td>
			<td class=tablebody1><a href="#" onclick="form1.code.value='<%=codeboxtext%>'">〖获取文字链接代码〗</a> <a href="#" onclick="form1.code.value='<%=codeboxlogo%>'">〖获取LOGO链接代码〗</a></td>
		</tr>
		<form name="form1">
		<tr> 
			<td colspan="2" class=tablebody1 align="center"><textarea name="code" rows=3 cols=100 style="border:1px Dashed" onMouseOver="form1.code.select()"></textarea></td>
		</tr>
		</form>
		<tr> 
			<td width="40%" class=tablebody1>联系站长：</td>
			<td class=tablebody1><a href="mailto:<%=forum_info(5)%>"><%=forum_info(5)%></a></td>
		</tr>
	</table><%
end if
yjsk_rst.close
set yjsk_rst=nothing
call activeonline()
call footer()
%>