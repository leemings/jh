<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>TD {
	FONT-SIZE: 12px; COLOR: #000000
}
</STYLE>

<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>
<BODY bgColor=#f9a20a leftMargin=0 topMargin=0>
<TABLE cellSpacing=0 cellPadding=0 width=738 border=0>
  <TBODY>
  <TR>
      <TD><A href="index.asp" target=_parent><IMG height=60 
      src="IMAGES/t1.gif" width=149 border=0></A><IMG height=60 
      src="images/newtop_02.gif" width=96><IMG height=60 
      src="images/newtop_03.gif" width=70><IMG height=60 
      src="images/newtop_04.gif" width=84><IMG height=60 
      src="images/newtop_05.gif" width=84><IMG height=60 
      src="images/newtop_06.gif" width=84><IMG height=60 
      src="images/newtop_07.gif" width=86><A 
      href="index.asp" target=_parent><IMG 
      height=60 alt=社区首页 src="images/newtop_08.gif" width=85 
  border=0></A></TD>
    </TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=738 border=0>
  <TBODY>
    <TR>
      <TD><A href="index.asp" target=_parent><IMG height=28 
      src="images/newtop_09.gif" width=149 border=0></A><IMG height=28 
      src="images/newtop_10.gif" width=96><A 
      href="index.asp?id=1" target=_parent><IMG 
      height=28 alt=推荐商品 src="IMAGES/top_1.gif" width=70 
      border=0></A><IMG height=28 src="images/newtop_12.gif" width=14><A 
      href="index.asp?id=2" target=_parent><IMG 
      height=28 alt=潮流服饰 src="IMAGES/top_2.gif" width=70 
      border=0></A><IMG 
      height=28 src="images/newtop_14.gif" width=14><A 
      href="index.asp?id=3" target=_parent><IMG 
      height=28 alt=化妆整形 src="IMAGES/top_3.gif" width=70 
      border=0></A><IMG height=28 src="images/newtop_16.gif" width=14><A 
      href="index.asp?id=4" target=_parent><IMG 
      height=28 alt=珠宝首饰 src="IMAGES/top_4.gif" width=70 
      border=0></A><IMG height=28 src="images/newtop_18.gif" width=16><A 
      href="index.asp?id=5" target=_parent><IMG 
      height=28 alt=精品场景 src="IMAGES/top_5.gif" width=70 
  border=0></A><IMG height=28 src="images/newtop_21.gif" width=15><IMG 
      height=28 src="images/newtop_14.gif" width=14><IMG height=28 src="images/newtop_21.gif" width=15><IMG height=28 src="images/newtop_16.gif" width=14><IMG height=28 src="images/newtop_16.gif" width=14></TD>
    </TR>
  </TBODY>
</TABLE>
<MAP name=Map3Map><AREA shape=RECT 
  target=_parent alt=返回首页 coords=23,1,142,11 
  href="index.asp"></MAP>
<TABLE height=19 cellSpacing=0 cellPadding=0 width="100%" 
background=IMAGES/top_<%=trim(Request.QueryString("id"))%>11.gif border=0>
  <TBODY>
  <TR>
    <TD width=245>&nbsp;</TD>
    <TD>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
<%
id=trim(Request.QueryString("id"))
sid=trim(Request.QueryString("sid"))
if id<>"" then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("showmdb")
	rs.Open "select * from fl where fl="& id , conn, 1,3
	do while Not RS.Eof
	Response.Write"<TD>·<A href=""main.asp?id="& rs("fl") &"&sid="& rs("fltype") &""" target=mainFrame><FONT color=#000000>"& rs("flname") &"</FONT></A></TD>"
	RS.MoveNext
	loop
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
%>        

</TR></TBODY></TABLE></TD></TR></TBODY></TABLE></BODY></HTML>
