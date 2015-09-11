<!-- #include file="setup.asp" -->
<%
if Request("message")<>"" then
top
%>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 社区信息</td>
</tr>
</table><br><SCRIPT>valigntop()</SCRIPT>
<table cellspacing="0" cellpadding="0" width="97%" align="center" border="0" class="a2"><tr><td height="106"><table cellspacing="1" cellpadding="6" width="100%" border="0"><tr><td width="100%" height="20" colspan="2" align="middle" class="a1">
<b>社区提示信息</b></td></tr><tr bgcolor="ffffff"><td valign="top" align="left" colspan="2" height="122"><b>&nbsp;</b><table cellspacing="0" cellpadding="5" width="100%" border="0"><tr><td width="83%" valign="top"><b>操作不成功的可能原因：</b><ul><%=Request("message")%></ul></td><td width="17%"><img height="97" src="images/err.gif" width="95"></td></tr></table></td></tr><tr align="middle" bgcolor="ffffff"><td valign="center" colspan="2" height="2"><input onclick="history.back(-1)" type="submit" value=" &lt;&lt; 返 回 上 一 页 " name="Submit"> </td></tr></table></td></tr></table>
<SCRIPT>valignbottom()</SCRIPT>
<%else%>
<script>alert('<%=Request("message2")%>');history.back();</script><script>window.close();</script>
<%
end if
htmlend
%>