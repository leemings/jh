<%@ LANGUAGE=VBScript codepage ="936" %>
<%response.expires=0
%><html>
<head>
<title>排行榜</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<%
const MaxPerPage=10
dim totalPut
dim CurrentPage
dim TotalPages
dim i,j
act=Request("act")
if act<>"道德" and act<>"总杀人" and act<>"女富翁" and act<>"男富翁" and act<>"财富排行" and act<>"金币" and act<>"会员金卡" and act<>"武功" and act<>"攻击" and act<>"防御" then
	act="道德"
end if
%>
<script language="JavaScript">
var chatbgcolor = parent.chatbgcolor;
var chatimage = parent.chatimage;
document.write("<body oncontextmenu=self.event.returnValue=false bgcolor=" + chatbgcolor + " background=" + chatimage + " bgproperties=fixed topMargin=10><center>");
function gourl(urldz)
{
location.href="KILLMAN.ASP?act="+urldz
}
</script>
<table border="1" width="98%" cellspacing="0" cellpadding="0" bgcolor="#336633" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
    <td width="100%"> 
      <div align="center"> 
        <select name="select" onChange='gourl(this.value);'>
          <option value="道德" <%if act="道德" then%>selected<%end if%>>江湖恶人</option>
          <option value="总杀人" <%if act="总杀人" then%>selected<%end if%>>总杀人数</option>
          <option value="女富翁" <%if act="女富翁" then%>selected<%end if%>>女 富 姐</option>
          <option value="男富翁" <%if act="男富翁" then%>selected<%end if%>>男 款 爷</option>
          <option value="财富排行" <%if act="财富排行" then%>selected<%end if%>>财富排行</option>
          <option value="金币" <%if act="金币" then%>selected<%end if%>>金币排行</option>
          <option value="会员金卡" <%if act="会员金卡" then%>selected<%end if%>>会员金卡</option>
          <option value="武功" <%if act="武功" then%>selected<%end if%>>武功排行</option>
          <option value="攻击" <%if act="攻击" then%>selected<%end if%>>攻击排行</option>
          <option value="防御" <%if act="防御" then%>selected<%end if%>>防御排行</option>
        </select>
      </div>
</td>
</tr>
</table>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
dim sql
dim rs
dim filename
sl=20
select case act
case "总杀人"
	rs.open "select top "& sl &"  姓名,"& act &" from [用户] order by 总杀人 desc",conn,1,1
case "女富翁"
	act="存款"
	rs.open "select top "& sl &"  姓名,"& act &" from [用户] where 性别='女' and grade<10 order by 存款 desc",conn,1,1
case "男富翁"
	act="存款"
	rs.open "select top "& sl &"  姓名,"& act &" from [用户] where 性别='男' and grade<10 order by 存款 desc",conn,1,1
case "财富排行"
	act="存款"
	rs.open "select top "& sl &"  姓名,"& act &" from [用户]  where grade<10 order by 存款 desc",conn,1,1
case "金币"
	rs.open "select top "& sl &"  姓名,"& act &" from [用户] where grade<10 order by 金币 desc",conn,1,1
case "会员金卡"
	rs.open "select top "& sl &"  姓名,"& act &" from [用户]  where grade<10 order by 会员金卡 desc",conn,1,1
case "武功"
	rs.open "select top "& sl &"  姓名,"& act &" from [用户]  where grade<10 order by 武功 desc",conn,1,1
case "攻击"
	rs.open "select top "& sl &"  姓名,"& act &" from [用户]  where grade<10 order by 攻击 desc",conn,1,1
case "防御"
	rs.open "select top "& sl &"  姓名,"& act &" from [用户]  where grade<10 order by 防御 desc",conn,1,1
case else
	rs.open "select top "& sl &" 姓名,"& act &" from [用户]  where grade<10 order by 道德 ",conn,1,1
end select
if rs.eof and rs.bof then
	response.write "<p align='center'>没有可排行的对象 </p>"
else
%>
<table border="1" cellspacing="0" width="98%" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="4" align="center">
  <tr bgcolor="#336633">
    <td align="center"><font color="#FFFFFF">人 名</font></td>
    <td align="center"><font color="#FFFFFF"><%=act%></font></td>
</tr>
<%filename=0
do while not rs.eof%>
<tr bgcolor="#CCCCCC" onmouseover="this.bgColor='#85C2E0';" onmouseout="this.bgColor='#CCCCCC';">
    <td align="center"><%=rs("姓名")%></td>
    <td align="center"><%=rs(act)%></td>
</tr>
<%
rs.movenext
filename=filename+1
if filename>sl then Exit Do
loop
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</body>
</html>