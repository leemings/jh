<%@ LANGUAGE=VBScript%>
<!--#include file="jhb.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if session("yx8_mhjh_inchat")<>"1" then %>
<script language="vbscript">
alert "请进入聊天室再进入股票！"
window.close()
</script>
<%Response.End
end if

Jname=username
sql="delete * from 交易 where 时间<date()"
conn.execute sql
sql="select * from 交易 where 帐号='" & Jname & "' "
set rs=conn.execute(sql)

%>
<table border=1 bgcolor="#bee0fc" align=center width=550 cellpadding="10" cellspacing="13">
	<option><p align=center><font size=5 color=#00ff00></font></p>
	<tr><td bgcolor=#cce8d6>
		<table border=0 cellpadding=1 cellspacing=1 bgcolor="#51a8ff" width="100%" >		</table>
		
		<P align=center><front size=3pt>经纪人今天为您代理的交易如下，扣除1%的佣金就是您在股市的收益</front></p>
		<table border=0 cellpadding=1 cellspacing=1 bgcolor="#51a8ff" width="100%" 
     >
	
		
		
        <P align=center></P>
        <TBODY>
			<tr bgcolor=#c4deff>
				<td>股票名称</td><td>买卖操作</td><td>交易数量</td><td>时间</td><td>收益</td>
			</tr>
<%DO UNTIL RS.EOF%>	
			<tr bgcolor=#f7f7f7>
				<td><%=rs("企业")%>
				<td><%=rs("操作")%></td>
				<td><%=rs("交易量")%></td>
				<td><%=rs("时间")%></td>
				<td><%=formatnumber(rs("收获"),0)%></td>
				
		    </tr>
<%
rs.movenext
loop
conn.close
%></p>

		</table>