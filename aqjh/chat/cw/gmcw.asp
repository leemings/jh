<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select cw from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
if isnull(rs("cw")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您还没有宠物，请去购买！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
zt=split(rs("cw"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：宠物数据出错，请重新购买！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
rs.close
rs.open "select i from b where a='"&zt(1)&"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：宠物数据出错，请重新购买！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
mypic=rs("i")
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
<title>宠物改名</title>
<script language="JavaScript">
function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}</script>
<style>
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../<%=chatimage%>" bgproperties="fixed" leftmargin="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<div align="center"><font color="#FFFFFF"><b>改 名 宠 物</b></font><br>
  <br>
  <br>
</div>
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%" >
    <form method="post" action="gmcwok.asp" id=form1 name=form1 target=d>
      <tr align="center"> 
        
      <td colspan="2"> 宠物原名 
        <input type="text" readonly name="oldname" size="10" maxlength="10" value="<%=zt(0)%>">
          <br>
        宠物新名 
        <input type="text" name="newname" size="10" maxlength="10">
          <br>
          <input type="submit" name="Submit" value="提交">
        </td>
      </tr>
    </form>
  </table>
  
<div align="center"><font color="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;<font size="2"><a href=javascript:history.go(-1)><font color=blue size=-1>返回上级</font></a><br>
  <br>
  <font color="#FF0000">警告:取的名字不雅，将删除你的ID！</font></font></font></div>
</body>
</html>