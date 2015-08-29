<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
if sjjh_name<>Application("sjjh_user") then 
	Response.Write "<script Language=Javascript>alert('提示：你不是正站长，你不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
cz=trim(request.querystring("cz"))
if cz<>"备份" then cz=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM sm where a='广告"&cz&"'",conn,2,2
%>
<html>
<head>
<title>聊天室顶部广告修改♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center">聊天室顶部广告修改(&lt;br&gt;为回车 支持html语言)
<p><br>
  <font size="+2" face="楷体_GB2312">如果资料修改错了,会造成一些想不到的后果，如果进入不了聊天室等，如果你修改错了请点<a href="ggsm.asp?cz=%B1%B8%B7%DD"><b><font color="#FF0000">这里</font></b></a>进行必要的修正，再点修改就可以恢复到初始化的状态 
  ，注意修改不要错了!</font>
<p align="center">(<font color="#0000FF">办理会员说明</font>) <br>

<form method=POST action="ggsmok.asp?cz=广告" name="form">
  <div align="center"> 广告宽度： 
    <input type="text" name="hydz" size="10" value="<%=rs("b")%>" maxlength="2">
    (设置为0则关闭广告)<br>
    <textarea name="hysm"  cols="90" rows="14"><%=rs("c")%></textarea>
    <br>
    <br>
    <br>
    <input type=submit value=确定顶部广告修改 name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
  </div>
</form>
<%rs.close
rs.open "SELECT * FROM sm where a='进入"&cz&"'",conn,2,2%>
<form method=POST action="ggsmok.asp?cz=进入" name="form">
  <div align="center"><font color="#0000FF"><b>进入江湖的广告设置</b></font><br>
    (非会员，每一条之间以;分号半角分隔)<br>
    <textarea name="hysm"  cols="90" rows="8"><%=rs("c")%></textarea>
    <br>
    (会员进入说明，只有一句话\n表示换行!) <br>
    <textarea name="hydz"  cols="90" rows="4"><%=rs("d")%></textarea>
    <br>
    <br>
    <br>
    <input type=submit value=确定进入广告修改 name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
  </div>
</form>
<%rs.close
rs.open "SELECT * FROM sm where a='滚动"&cz&"'",conn,2,2%>
<form method=POST action="ggsmok.asp?cz=滚动" name="form">
  <div align="center"><font color="#0000FF"><b>聊天室上面滚动广告设置</b></font><br>
    每一条之间以;分号半角分隔<br>
    <textarea name="hysm"  cols="90" rows="12"><%=rs("c")%></textarea>
    <br>
    江湖博士格式如下：问1|答案1;问2|答案2;<br>
    <textarea name="hydz"  cols="90" rows="12"><%=rs("d")%></textarea>
    <br>
    <br>
    <br>
    <input type=submit value=确定滚动广告修改 name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
  </div>
</form>
<%rs.close
set rs=nothing
conn.close
set conn=nothing%> </body>
</html>
