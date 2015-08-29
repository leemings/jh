<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="jmconfig.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=session("yx8_mhjh_username")
sjjh_grade=session("yx8_mhjh_usergrade")
sjjh_jhdj=session("yx8_mhjh_usergrade")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"

mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("zgzm.asp")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
set rs1=conn.execute("select * from lb")
%>
<html>
<head>
<title>周公解梦</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<script language="JavaScript">
<!--
function check()
{
	var locat=document.searchs.locat.value;
	if(locat=="" || locat==null)
	{
		window.alert("请输入要搜索的关键字！");
		document.searchs.locat.focus;
		return false;
	}
else
	{return true;}
}
// -->
</script>
<body bgcolor="#FFFFFF" text="#000000" style="font-size: 8pt; background-color: #FFCCFF; border: 1 solid #F5A0E2">
<p align="center"><font size="4" face="楷体_GB2312"><b class="3d"><font size="5">快乐江湖 
  之 周公解梦</font></b></font></p>
<div align="center"><table border="0" width="85%" cellpadding="0" cellspacing="0">
  <tr><form method="POST" action="search.asp" target=_blank name="searchs">
        <td width="100%"> 
          <p align="center">输入关键字（支持模糊查询） 
            <input type="text" name="locat" size="35" class="input">
            <input type="button" value="搜索" name="B1" style="border-style: solid; border-width: 1" class="input" onclick="javascript:if (check()==true){this.document.searchs.submit('search.asp')}">
            <input type="reset" value="清除" name="B2" class="input">
          </p>
      
    </td></form>
  </tr>
</table>
<br>

  <table width="716" cellspacing="0" cellpadding="0" height="62" bordercolor="#FF99FF" bordercolorlight="#FF99FF" style="border: 1 solid #FF99FF" bordercolordark="#FF99FF">
    <tr bgcolor="#CC66CC"> 
      <td width="93" height="29"> 
        <div align="center"><b><font size="2">类别</font></b></div>
    </td>
      <td width="621" height="29"><b><font size="2">解梦内容(每条收取 银两:<%=jmyin%>，金币:<%=jmjb%>)</font></b></td> 
  </tr>
  <%do while not(rs1.eof or rs1.bof)%>
  <tr> 
      <td width="93" height="24" valign="middle" bordercolor="#FF99FF" bordercolorlight="#FF99FF" bordercolordark="#FF99FF" style="border-right: 1 solid #FF99FF; border-bottom: 1 solid #FF99FF"> 
        <div align="center"><font size="2"><%=rs1("lb")%></font></div>
    </td>
    <%rs.open "select id,a,b from zgjm where a="&rs1("id"),conn%>
      <td width="621" height="24" style="border-bottom: 1 solid #FF99FF"> <font size="2"> 
        <p style="line-height: 150%"> 
        <%if rs.eof or rs.bof then
			response.write "<font color=red><b>此类别暂无任何内容！</b></font>"
		else
			do while not (rs.eof or rs.bof)
				response.write "<a href='#' onClick=window.open('jm.asp?id="&rs("id")&"','','scrollbars=yes,resizable=yes,width=640,height=420')>"&rs("b")&"</a> "
				rs.movenext
			loop
			rs.close
		end if%>
      </font></td>
  </tr>
  <%rs1.movenext
  loop
  rs1.close
  set rs1=nothing
  set rs=nothing
  conn.close
  set conn=nothing%>
</table>
  <p><font size="2">辽西江湖制作</font></p>
</div>
</body>
</html>
