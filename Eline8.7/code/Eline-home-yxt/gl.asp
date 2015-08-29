<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
sjjh_name=sjjh_name
grade=Int(sjjh_grade)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
dim tmprs
tmprs=conn.execute("Select count(*) As 数量 from 用户 where times<3 and CDate(lasttime)<now()-10")
musers=tmprs("数量")
set tmprs=nothing
if isnull(musers) then musers=0
user1=musers
tmprs=conn.execute("Select count(*) As 数量 from 用户 where allvalue<=5 and CDate(lasttime)<now()-10")
musers=tmprs("数量")
set tmprs=nothing
if isnull(musers) then musers=0
user2=musers
tmprs=conn.execute("Select count(*) As 数量 from 用户 where CDate(lasttime)<now()-30")
musers=tmprs("数量")
set tmprs=nothing
if isnull(musers) then musers=0
user3=musers
%><title><%=Application("sjjh_chatroomname")%>数据库维护程序</title>
<body background="../JHIMG/Bk_hc3w.gif">
<div align="center">
<p>主用户数据库：</p>
<table width="500" border="1" bordercolor="#000000" cellspacing="0" cellpadding="1">
<tr>
<td width="85%">10天之前,仅登陆3次的用户有：<font color="#0000FF"><b><%=user1%>个</b></font></td>
<td width="15%">
<div align="center"><a href="gl1.asp">删除</a></div>
</td>
</tr>
<tr>
<td width="85%">10天之前泡点数allvalue小于5的用户有：<b><font color="#0000FF"><%=user2%>个</font></b></td>
<td width="15%">
<div align="center"><a href="gl2.asp">删除</a></div>
</td>
</tr>
<tr>
<td width="85%">30天从未登陆的用户有：<font color="#0000FF"><b><%=user3%>个</b></font></td>
<td width="15%">
<div align="center"><a href="gl3.asp">删除</a></div>
</td>
</tr>
<tr>
<td width="85%">赌博数据修正，如果无法开局提供没有到请使用此修正！</td>
<td width="15%">
<div align="center"><a href="gl5.asp">修正</a></div>
</td>
</tr>

</table>
  <p align="center"><BR>
    此套程序由<%=Application("sjjh_chatroomname")%>小小宇完成！<br>
    在这里你可以查看到当前数据库的使用情况，<br>
    选择删除将可以把这一些用户删除掉，删除之后选择数据库压缩操作！<br>
    在我的江湖上我就用这一些方法成功把一个11MB的数据库压缩成4MB大小！<br>
    在此操作前建议备份原始文件，如果因为程序操作造成的后果我们不负任何责任！<br>
    对于:记录数据库、BBS数据库如果特别大的可以选择用原始文件覆盖安即可，<br>
    建议对这些文件备份<BR>
  </p>
</div>