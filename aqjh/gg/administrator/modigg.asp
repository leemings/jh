<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../../error.asp?id=439"
id=clng(Request("id"))
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../gg.asa")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
rs.open "select * from 公告 where id="&id,conn,1,2
if rs.EOF or rs.BOF then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "ggadmin.asp"
	Response.End
end if
gg=rs("内容")
bt=rs("标题")
fbsj=rs("时间")
ck=rs("查看次数")
dqlb=rs("所属类别")
rs.close
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../hc3w_Admin/setup.css">
<title>爱情江湖物品修改</title></head>

<body text="#000000" background="../../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center"><font size="2" color="#000000">爱情江湖首页站长公告修改程序！<br>
  </font></p>
<form method="post" action="modiggok.asp">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="560">
    <tr> 
      <td width="90"> <font color="#FFFFFF" size="2">ＩＤ：</font></td>
      <td width="472"><font color="#FFFFFF" size="2"> 
        <input type="text" name="id" readonly value="<%=id%>" size="5">
        </font></td>
    </tr>
    <tr> 
      <td width="90"> <font color="#FFFFFF" size="2">标题：</font></td>
      <td width="472"><font color="#FFFFFF" size="2"> 
        <input type="text" name="bt"
value="<%=bt%>" size="40">
        </font></td>
    </tr>
    <tr> 
      <td width="90"> <font color="#FFFFFF" size="2">所属类别：</font></td>
      <%rs.open "select * from l",conn%>
	  <td width="472"><font color="#FFFFFF" size="2">
		<select name="lb">
		  <%do while not(rs.bof or rs.eof)
		  z=rs("id")&"|"&rs("lb")%>
          <option value="<%=z%>" <%if dqlb=rs("id") then%>selected<%end if%>><font color="#FFFFFF" size="2"><%=rs("lb")%></font></option>
          <%rs.movenext
		  loop
		  %>
        </select>
        </font></td>
		<%rs.close
		  set rs=nothing
		  conn.close%>
    </tr><tr> 
      <td width="90"> <font color="#FFFFFF" size="2">发表时间：</font></td>
      <td width="472"><font color="#FFFFFF" size="2"> 
        <input type="text" name="sj" value="<%=fbsj%>" size="30">
        </font></td>
    </tr>
    <tr> 
      <td width="90"> <font color="#FFFFFF" size="2">点击次数：</font></td>
      <td width="472"><font color="#FFFFFF" size="2"> 
        <input type="text" name="ck"
value="<%=ck%>" size="10">
        </font></td>
    </tr>
    <tr> 
      <td width="90" valign="top"> <font color="#FFFFFF" size="2"><br>
        内容：</font></td>
      <td width="472"><font color="#FFFFFF" size="2"> 
        <textarea name="gg" cols="65" rows="10"><%=gg%></textarea>
        </font></td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center">
          <input type="submit" value="确 定">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.document.location.href='ggadmin.asp'" value="返 回" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
<div align="center">本程序由爱情江湖总站<font color="#0000FF">永不放弃</font>制作</div>
