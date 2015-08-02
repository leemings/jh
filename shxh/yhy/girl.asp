
<%
nickname=Session("Ba_jxqy_username")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("Ba_jxqy_connstr")
conn.open connstr
sql="select * from 用户 where 姓名='" & nickname & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	  Response.Redirect "../error.asp?id=016"
	conn.close
	response.end
else
dj=rs("等级")
yl=rs("银两")
xingbie=rs("性别")
if dj<3 then
	response.write "你还是江湖小辈，就想来这种地方"
	conn.close
	response.end
else
if xingbie="女" then
	response.write "怡红院的姑娘不接待女的，请止步！"
	conn.close
	response.end

else %>
<!--#include file="jiu.asp"-->
<% id=request("id")
sql="select * from 名妓 where ID=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
response.write "你有没有搞错呀，本院哪有这个姑娘呀，是不是想姑娘想疯了"
	connt.close
	response.end
else
mingji=rs("姓名")
meimao=rs("美貌度")

if yl<meimao*1.5  then
	response.write "没钱也想来这种高级的地方呀，请止步！"
            conn.close
	response.end
else
sql="update 用户 set 内力=内力+"&meimao&"/2,体力=体力-100 where 姓名='"&nickname&"' "
conn.execute sql
sql1="update 用户 set 银两=银两-"&meimao&"*1.5 where 姓名='"&nickname&"' "
conn.execute sql1
connt.close
conn.close


end if
end if
end if
end if
end if%>
<html>
<title>海阔天空-歌舞宴会</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
.unnamed1 {  font-size: 9pt}
-->
</style>

<body bgcolor="#FFB366">
<p>　</p>
<table width="52%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="17" bgcolor="#996633" align="center">　</td>
  </tr>
  <tr bgcolor="#66FF66"> 
    <td align="center" height="378" bgcolor="#FFCC66"> 
      <p><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="550" height="400">
          <param name=movie value="girl.swf">
          <param name=quality value=high>
          <embed src="girl.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="550" height="400">
          </embed></object><font> </font></p>
    </td>
  </tr>
  <tr bgcolor="#0033CC"> 
    <td align="center" height="26" class="unnamed1" bgcolor="#FFCC66"><b></b></td>
  </tr>
</table>
<p>　</p>
</body>
</html>