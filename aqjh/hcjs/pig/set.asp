<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from 猪圈设置"
rs.open sql,connt,1,1
if rs.recordcount<>0 then
bm_start=rs("报名开始")
bm_end=rs("报名结束")
tp_start=rs("投票开始")
tp_end=rs("投票结束")
tp_dj=rs("等级")
end if
%>
<HTML><HEAD>
<TITLE>〖养猪大赛〗</TITLE>
<link rel="stylesheet" href="../../css.css">
</HEAD>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<table width=500 border="0" align=center cellspacing="0" bordercolor="#FF0000">
        <tr bordercolor="#FFFFFF">
          <td align=center><font color=red><b>养 猪 大 赛 设 置</b></font><BR><BR></td>
        </tr>
      </table>
      <table width=500 border="0" align=center cellspacing="0" bordercolorlight="#800080"
bordercolordark="#FFFFFF">
<tr><td><center>
<p>格式为“年-月-日 时:分:秒”，例：“2000-02-17 08:00:00”（全部为半角字符）。</p>
<p><a href=cl.asp><font color=red>清理数据</font></a>(注意：清理后将不能恢复，请务必小心)</p>
</blockquote>
<table border="1" align="center" bgcolor="#E0E0E0" cellpadding="4" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellspacing="0">
<form method="post" action="setok.asp">
<tr>
<td>投票所需等级：</td>
<td>
<input type="text" name="tp_dj" maxlength="19" size="19" value="<%=tp_dj%>">
</td>
</tr>
<tr>
<td>开始报名时间：</td>
<td>
<input type="text" name="bm_start" maxlength="19" size="19" value="<%=bm_start%>">
</td>
</tr>
<tr>
<td>终止报名时间：</td>
<td>
<input type="text" name="bm_end" maxlength="19" size="19" value="<%=bm_end%>">
</td>
</tr><tr>
<td>开始投票时间：</td>
<td>
<input type="text" name="tp_start" maxlength="19" size="19" value="<%=tp_start%>">
</td>
</tr>
<tr>
<td>终止投票时间：</td>
<td>
<input type="text" name="tp_end" maxlength="19" size="19" value="<%=tp_end%>">
</td>
</tr>
<tr>
<td align="right">选择操作：</td>
<td align="center">
<input type="submit" name="Submit" value="更改">
<input type="button" value="放弃" onclick=javascript:history.go(-1)>
</td>
</form>
</table>
</td></tr></table></body></html>