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
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
end if
if rs("泡豆点数")<1000 then
	mess="你现在还没有豆子，只有:<font color=red>"&rs("泡豆点数")&"</font>点"
else
	dousl=int(rs("泡豆点数")/1000)
	conn.execute "update 用户 set 泡豆点数=泡豆点数-"& dousl*1000 &",会员金卡=会员金卡+"& 2*dousl &" where 姓名='"&aqjh_name&"'"
	mess="卖豆子："&dousl&"个，会员费增加："&dousl*2&"元，还剩:"&rs("泡豆点数")-dousl*1000&"点！"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<link rel="stylesheet" href="../../css.css">
<title>卖豆子</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bg.gif">
<table border="0" cellspacing="0" cellpadding="0" width="259" align="center">
  <tr>
<td height="81" valign="top">
      <div align="center"><font color="#000000"><b><font color=blue><%=aqjh_name%></font>欢迎光临自由市场卖豆</b></font></div>
      <table width="280" align="center">
        <tr> 
            <td> 
          <tr> 
            
          <td valign="top" height="11" > 
            <div align="center">豆子简介</div>
            </td>
          </tr>
          <tr> 
            
          <td valign="top" height="150" > 
            <p>豆子是由你泡点时攒下来了,与泡点另算。每泡够1000点，计算机系统自动给你一个豆，当你选择暴豆后，杀伤力为平时3倍，可以持续1小时时间。如果你不想使用，可以在这里卖了。一个豆子=2元会员费用，你可以用卖豆子的会员费购买只有会员才有的卡片！</p>
<%=mess%><br><br>
            </td>
          </tr>
        </table>

</td>
</tr>
</table>
<div align="center"><font color="#00FF66"><font color="#0000FF">≮爱情江湖总站≯</font></font>
</div>
</body>
</html>