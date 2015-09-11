<html>
<head>
<LINK href="../css.css" rel=stylesheet>
<script language="JavaScript">
<!--
 
if (window.Event) 
  document.captureEvents(Event.MOUSEUP); 
 
function nocontextmenu() 
{
 event.cancelBubble = true
 event.returnValue = false;
 
 return false;
}
 
function norightclick(e) 
{
 if (window.Event) 
 {
  if (e.which == 2 || e.which == 3)
   return false;
 }
 else
  if (event.button == 2 || event.button == 3)
  {
   event.cancelBubble = true
   event.returnValue = false;
   return false;
  }
 
}
 
document.oncontextmenu = nocontextmenu;  // for IE5+
document.onmousedown = norightclick;  // for all others
//-->
</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=000000 leftMargin="0" topMargin="0" MARGINHEIGHT="0" MARGINWIDTH="0">
<table width=510 align=center height=256 background="imgg/bj.jpg">
<tr>
		<td width=30%>&nbsp;</td>
		<td align=center height=256 width=70%><br><br><br>&nbsp;&nbsp;
          <font color=ffffff><font color=ffffff>这里只对爱情茶壶任务的侠客提供服务</font>
<br><br>
<a href="rwu1/index.asp" onClick="javascript:s()"><font color=red>执行爱情茶壶任务</font></a></td>
	</tr>
</table>
</body>
</html>                             
</html>
