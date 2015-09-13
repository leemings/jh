<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>快乐江湖购买</title>
<script language="JavaScript" type="text/JavaScript">
window.moveTo(100,50);
function show(myshow){Layer1.style.visibility="visible";msgshow.innerHTML=myshow;}
function hidden(){Layer1.style.visibility="hidden";}
document.onmousedown=click;
function click(){if(event.button==2){alert("欢迎来到快乐江湖！");}}
</script>
<LINK href="style.css" rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0">
<table width="75%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="635" height="397"><img src="pic/bk.jpg" width="635" height="397" border="0" usemap="#Map"></td>
  </tr>
</table>
<div id="Layer1" name="Layer1" style="position:absolute; width:212px; height:88px; z-index:1; left: 410px; top: 5px; background-image: url(pic/show.jpg); layer-background-image: url(pic/show.jpg); border: 1px none #000000; visibility: hidden;"> 
  <table width="95%" height="89" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="16">&nbsp;</td>
    </tr>
    <tr>
      <td height="55" align="left" valign="top" id="msgshow"> </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
</div>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from [housetype] where ht_序号=5",conn,1,1
%>
<map name="Map">
  <area shape="rect" coords="220,53,268,126" href="buy.asp?id=5&name=<%=trim(rs("ht_小区名"))%>&tj=<%=trim(rs("ht_1级条件"))%>" onmouseover="show('<%=trim(rs("ht_小区名"))%>:<%=trim(rs("ht_小区说明"))%><br>居民:<%=rs("ht_居民")%>人 资产:<%=rs("ht_小区资产")%>万<br>区长:<%=rs("ht_小区区长")%>');" onmouseout="hidden();">
<%rs.close
rs.open "Select * from [housetype] where ht_序号=1",conn,1,1%>  
  <area shape="rect" coords="82,200,160,255" href="buy.asp?id=1&name=<%=trim(rs("ht_小区名"))%>&tj=<%=trim(rs("ht_1级条件"))%>" onmouseover="show('<%=trim(rs("ht_小区名"))%>:<%=trim(rs("ht_小区说明"))%><br>居民:<%=rs("ht_居民")%>人 资产:<%=rs("ht_小区资产")%><br>区长:<%=rs("ht_小区区长")%>');" onmouseout="hidden();" >
<%rs.close
rs.open "Select * from [housetype] where ht_序号=4",conn,1,1%>  
  <area shape="rect" coords="171,288,255,350" href="buy.asp?id=4&name=<%=trim(rs("ht_小区名"))%>&tj=<%=trim(rs("ht_1级条件"))%>" onmouseover="show('<%=trim(rs("ht_小区名"))%>:<%=trim(rs("ht_小区说明"))%><br>居民:<%=rs("ht_居民")%>人 资产:<%=rs("ht_小区资产")%><br>区长:<%=rs("ht_小区区长")%>');" onmouseout="hidden();">
<%rs.close
rs.open "Select * from [housetype] where ht_序号=0",conn,1,1%>  
  <area shape="rect" coords="421,285,508,346" href="buy.asp?id=0&name=<%=trim(rs("ht_小区名"))%>&tj=<%=trim(rs("ht_1级条件"))%>" onmouseover="show('<%=trim(rs("ht_小区名"))%>:<%=trim(rs("ht_小区说明"))%><br>居民:<%=rs("ht_居民")%>人 资产:<%=rs("ht_小区资产")%><br>区长:<%=rs("ht_小区区长")%>');" onmouseout="hidden();">
<%rs.close
rs.open "Select * from [housetype] where ht_序号=3",conn,1,1%>  
  <area shape="rect" coords="528,226,590,279" href="buy.asp?id=3&name=<%=trim(rs("ht_小区名"))%>&tj=<%=trim(rs("ht_1级条件"))%>" onmouseover="show('<%=trim(rs("ht_小区名"))%>:<%=trim(rs("ht_小区说明"))%><br>居民:<%=rs("ht_居民")%>人 资产:<%=rs("ht_小区资产")%><br>区长:<%=rs("ht_小区区长")%>');" onmouseout="hidden();">
<%rs.close
rs.open "Select * from [housetype] where ht_序号=2",conn,1,1%>  
  <area shape="rect" coords="410,145,503,208" href="buy.asp?id=2&name=<%=trim(rs("ht_小区名"))%>&tj=<%=trim(rs("ht_1级条件"))%>" onmouseover="show('<%=trim(rs("ht_小区名"))%>:<%=trim(rs("ht_小区说明"))%><br>居民:<%=rs("ht_居民")%>人 资产:<%=rs("ht_小区资产")%><br>区长:<%=rs("ht_小区区长")%>');" onmouseout="hidden();">
  <area shape="rect" coords="295,152,355,233" href="#" onmouseover="show('欢迎使用快乐江湖，在这里你可以购买属于自己的房屋！<br>有了房屋就可以取妻生子了．');" onmouseout="hidden();">
</map>
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</body>
</html>
