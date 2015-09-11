<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<HTML>
<HEAD>
<TITLE>造化三国</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content=zh-cn http-equiv=Content-Language>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<META content=FrontPage.Editor.Document name=ProgId>
<STYLE type=text/css>
</STYLE>
<link rel="stylesheet" href="dg/Setup.css">
</HEAD>
<BODY oncontextmenu=self.event.returnValue=false bgColor=#000000 text=#ffffff>
    <table width="560" height="12">
    <TD height="289" width="15"></TD>
    <TD rowspan="3" height="8" width="615">
      <DIV align=center>
      <CENTER>
      <TABLE border=0 cellPadding=0 cellSpacing=0 height=1 width=594>
        <TBODY>
        <TR>
          <TD colSpan=4 height=1 width=589>&nbsp;&nbsp; 
            这里聚集的都是一些三国好汉，看造化了。 
          <TD height=1 width=11 rowspan="3">     
           </TD>    
        </TR>    
        <TR>    
          <TD align=middle height=138 width=145><a href='#' title='我要打架' onClick="window.open('meiyong1/meiyong.asp?my=<%=Session("wyjh_nickname")%>','youyi','scrollbars=yes,resizable=yes,width=380,height=100')"><IMG border=0 src="meiyong1/3-10.jpg"></A></TD>    
          <TD align=middle height=138 width=145><a href='#' title='我要打架' onClick="window.open('meiyong1/meiyong.asp?my=<%=Session("wyjh_nickname")%>','youyi','scrollbars=yes,resizable=yes,width=380,height=100')"><IMG border=0 src="meiyong1/3-11.jpg"></A></TD>    
          <TD align=middle height=138 width=146><a href='#' title='我要打架' onClick="window.open('meiyong1/meiyong.asp?my=<%=Session("wyjh_nickname")%>','youyi','scrollbars=yes,resizable=yes,width=380,height=100')"><IMG border=0 src="meiyong1/3-12.jpg"></A></TD>    
          <TD align=middle height=138 width=146><a href='#' title='我要打架' onClick="window.open('meiyong1/meiyong.asp?my=<%=Session("wyjh_nickname")%>','youyi','scrollbars=yes,resizable=yes,width=380,height=100')"><IMG border=0 src="meiyong1/3-13.jpg"></A></TD>    
        </TR>    
        <TR>    
          <TD align=middle height=1 vAlign=top width=145><a href='#' title='我要打架' onClick="window.open('meiyong1/meiyong.asp?my=<%=Session("wyjh_nickname")%>','youyi','scrollbars=yes,resizable=yes,width=380,height=100')"><IMG border=0 src="meiyong1/3-14.jpg"></A>    
          </TD>    
          <TD align=middle height=1 vAlign=top width=145><a href='#' title='我要打架' onClick="window.open('meiyong1/meiyong.asp?my=<%=Session("wyjh_nickname")%>','youyi','scrollbars=yes,resizable=yes,width=380,height=100')"><IMG border=0 src="meiyong1/3-15.jpg"></A></TD>    
          <TD align=middle height=1 vAlign=top width=146><a href='#' title='我要打架' onClick="window.open('meiyong1/meiyong.asp?my=<%=Session("wyjh_nickname")%>','youyi','scrollbars=yes,resizable=yes,width=380,height=100')"><IMG border=0 src="meiyong1/3-16.jpg"></A>    
          </TD>    
          <TD align=middle height=1 vAlign=top width=146><a href='#' title='我要打架' onClick="window.open('meiyong1/meiyong.asp?my=<%=Session("wyjh_nickname")%>','youyi','scrollbars=yes,resizable=yes,width=380,height=100')"><IMG border=0 src="meiyong1/3-17.jpg"></A></TD>
        </TR>    
        <TR>    
          <TD align=middle height=1 vAlign=top width=145><a href='#' title='我要打架' onClick="window.open('meiyong1/meiyong.asp?my=<%=Session("wyjh_nickname")%>','youyi','scrollbars=yes,resizable=yes,width=380,height=100')"><IMG border=0 src="meiyong1/3-18.jpg"></A>    
          </TD>    
          <TD align=middle height=1 vAlign=top width=145><a href='#' title='我要打架' onClick="window.open('meiyong1/meiyong.asp?my=<%=Session("wyjh_nickname")%>','youyi','scrollbars=yes,resizable=yes,width=380,height=100')"><IMG border=0 src="meiyong1/3-19.jpg"></A></TD>    
          <TD align=middle height=1 vAlign=top width=146><a href='#' title='我要打架' onClick="window.open('meiyong1/meiyong.asp?my=<%=Session("wyjh_nickname")%>','youyi','scrollbars=yes,resizable=yes,width=380,height=100')"><IMG border=0 src="meiyong1/3-20.jpg"></A>    
          </TD>    
          <TD align=middle height=1 vAlign=top width=146><a href='#' title='我要打架' onClick="window.open('meiyong1/meiyong.asp?my=<%=Session("wyjh_nickname")%>','youyi','scrollbars=yes,resizable=yes,width=380,height=100')"><IMG border=0 src="meiyong1/3-21.jpg"></A></TD></TR></CENTER>    
       </table>    
    </table>    
  </BODY>  
 </html>  
  
  
  
  
  
  
  
  




