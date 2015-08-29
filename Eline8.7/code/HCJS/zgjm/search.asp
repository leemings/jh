<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="jmconfig.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
gjz=trim(request.form("locat"))
if gjz="" or gjz=null then
	response.write "<script language=javascript>{alert('错误：请输入关键字！');window.close();}</script>"
	response.end
end if
gjz=replace(gjz,"<","")
gjz=replace(gjz,">","")
gjz=replace(gjz,"&lt;","")
gjz=replace(gjz,"&gt;","")
gjz=replace(gjz,"o","")
gjz=replace(gjz,"O","")
gjz=replace(gjz,"r","")
gjz=replace(gjz,"R","")
gjz=replace(gjz,"union","")
gjz=replace(gjz,"'","")
gjz=replace(gjz,chr(34),"")
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("e_zgzm.asp")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
set tjzs=conn.execute("select count(b) as jjj from zgjm where instr(b,'"&gjz&"')<>0 or instr(c,'"&gjz&"')<>0")
zs=tjzs("jjj")
set tjzs=nothing
rs.open "select id,a,b from zgjm where instr(b,'"&gjz&"')<>0 or instr(c,'"&gjz&"')<>0 order by a",conn,1,2
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
<p align="center"><font size="4" face="楷体_GB2312"><b class="3d"><font size="5">鬼傲江湖 
  之 周公解梦</font></b></font></p>
<div align="center"><table border="0" width="85%" cellpadding="0" cellspacing="0">
  <tr><form method="POST" action="search.asp" name="searchs">
        <td width="100%"> 
          <p align="center">输入关键字（支持模糊查询）
            <input type="text" name="locat" size="35" class="input">
            <input type="button" value="搜索" name="B1" style="border-style: solid; border-width: 1" class="input" onclick="javascript:if (check()==true){this.document.searchs.submit('search.asp')}">
            <input type="reset" value="清除" name="B2" class="input">
          </p>
    	</td>
	  </form>
  </tr>
</table>
<br>
  <table width="716" cellspacing="0" cellpadding="0" height="62" bordercolor="#FF99FF" bordercolorlight="#FF99FF" style="border: 1 solid #FF99FF" bordercolordark="#FF99FF">
    <tr bgcolor="#CC66CC"> 
      <td colspan="2" height="29"> <b><font size="2">搜索：</font><font color="#00FFFF"><%=gjz%></font></b>，内容或标题中有<font color="#00FFFF"><%=gjz%></font>字符的共搜索到 
        <font color="#00FFFF"><b><%=zs%></b></font> 条，查看每条收取 <font color="#00FFFF">银两：<%=jmyin%>金币：<%=jmjb%></font></td>
    </tr>
<%if rs.eof or rs.bof then%>
	<tr><td align="center" colspan="2" height=29><br><br><br><font color=red>对不起，没有找到任何符合条件的记录！</font></td></tr>
<%else
	do while not(rs.eof or rs.bof)
		a=rs("a")
		set rs1=conn.execute("select * from lb where id="&a)
		if rs1.eof or rs1.bof then
			fl="未分类"
		else
			fl=rs1("lb")
		end if
		rs1.close%>
	<tr>
      <td width="93" height="24" valign="middle" bordercolor="#FF99FF" bordercolorlight="#FF99FF" bordercolordark="#FF99FF" style="border-right: 1 solid #FF99FF; border-bottom: 1 solid #FF99FF">
        <div align="center"><font size="2"><%=fl%></font></div>
      </td>
	  <td width="621" height="24" style="border-bottom: 1 solid #FF99FF"> <font size="2">
	  <p style="line-height: 150%">
	  <%n=true
	  do while n=true%>
          <%if not rs.eof then
			  b=rs("a")
			  if a=b then
			  	jmsay=replace(rs("b"),gjz,"<font color=red>"&gjz&"</font>")	
				response.write "<a href='#' onClick=window.open('jm.asp?id="&rs("id")&"','','scrollbars=yes,resizable=yes,width=640,height=420')>"&jmsay&"</a> "
			  	rs.movenext
			  else
				n=false
				rs.moveprevious
			  end if
			else
				rs.moveprevious
				n=false
			end if
	  loop
	  %>
	</font>
		</td></tr>
	<%rs.movenext
	loop
end if
rs.close
set rs=nothing
set rs1=nothing
conn.close
set conn=nothing
%>
  </table>
  <p><font size="2">『E线江湖』&#8482; 2003-2004</font></p>
</div>
</body>
</html>
