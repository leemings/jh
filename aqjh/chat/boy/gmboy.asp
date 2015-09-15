<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select boy,boysex,w1 from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
myhua=rs("w1")
if isnull(rs("boy")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您还没有小孩，赶快生育一个吧！');</script>"
	response.end
end if
zt=split(rs("boy"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：小孩数据出错，请重新购买！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if

mypic=rs("boysex")
rs.close
set rs=nothing
conn.close
set conn=nothing
if DateDiff("h",zt(7),now())<5 then
	zg=clng(zt(6))
else
	zg=0
end if
%>
<html>
<head>
<title>小孩改名</title>
<script language="JavaScript">
function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}</script>
<style>
body{
CURSOR: url('shubiao.ani');
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff}
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#006699" bgproperties="fixed" leftmargin="0" topmargin="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" background="../<%=chatimage%>">
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
    <td width="100%" height="28"> <div align="center"><font color="#ffffff" size="2">小孩改名</font></div></td>
  </tr>
</table>
<table background="../../bg2.gif" border="1" cellspacing="0" width="140" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="4" align="center">
    <form method="post" action="gmcwok.asp" id=form1 name=form1 target=d>
      <tr align="center"> 
        
      <td colspan="2"> 原名 
        <input type="text" readonly name="oldname" style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #009933 1px solid; BORDER-LEFT: #009933 1px solid; BORDER-RIGHT: #009933 1px solid; BORDER-TOP: #009933 1px solid; COLOR: black; FONT-SIZE: 9pt" size="10" maxlength="10" value="<%=zt(0)%>">
          <br>
        新名 
        <input type="text" name="newname" style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #009933 1px solid; BORDER-LEFT: #009933 1px solid; BORDER-RIGHT: #009933 1px solid; BORDER-TOP: #009933 1px solid; COLOR: black; FONT-SIZE: 9pt" size="10" maxlength="10">
          <br><br>
          <input type="submit" name="Submit" value="提交" style="BACKGROUND-COLOR: #009933; BORDER-BOTTOM: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; CURSOR: hand; FONT-FAMILY: '宋体'; COLOR: #FFFFFF; FONT-SIZE: 9pt; HEIGHT: 18px">
        </td>
      </tr>
    </form>
  </table>
  <br>
<div align="center">&nbsp;&nbsp;<font color="#FFFFFF"><font size="2"><a href=javascript:history.go(-1)><font color=#ffffff size=-1>返回上级</font></a><br>
  <br>
  <font color="#FF0000">警告:如果取的名字不雅，将删除你的ID！</font></font></font></div>
</body>
</html>