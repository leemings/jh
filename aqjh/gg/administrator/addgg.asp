<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../../error.asp?id=439"
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../hc3w_Admin/setup.css">
<title>添加物品</title></head>

<body text="#000000" background="../../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center"><font size="2" color="#000000">发布新公告</font></p>
<form method="post" action="addggok.asp">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="541">
    <tr> 
      <td colspan="2" height="18"><font color="#FFFF00" size="-1">提示:如无内容请添写为&quot;无&quot;</font></td>
    </tr>
    <tr> 
      <td width="49"> 
        <div align="center"><font color="#FFFFFF" size="2">标题：</font></div>
      </td>
      <td width="483"><font color="#FFFFFF" size="2"> 
        <input type="text" name="bt" size="40">
        </font></td>
    </tr>
	<tr> 
      <td width="49"> 
        <div align="center"><font color="#FFFFFF" size="2">标题：</font></div>
      </td>
	  <%mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../gg.asa")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
rs.open "select * from l",conn
%>
      <td width="483"><font color="#FFFFFF" size="2">
        <select name="lb">
          <option value="0|未分类" selected>指定类别</option>
          <%do while not(rs.bof or rs.eof)
		  z=rs("id")&"|"&rs("lb")%>
          <option value="<%=z%>"><%=rs("lb")%></option>
          <%rs.movenext
		  loop
		  %>
        </select>
        </font></td><%rs.close
		  set rs=nothing
		  conn.close%>
    </tr>
    <tr> 
      <td width="49" valign="top"> 
        <div align="center"><font color="#FFFFFF" size="2"><br>
          内容：</font></div>
      </td>
      <td width="483"><font color="#FFFFFF" size="2"> 
        <textarea name="gg" cols="65" rows="10"></textarea>
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
<div align="center">本程序由快乐江湖总站<font color="#0000FF">回首当年</font>制作</div>
