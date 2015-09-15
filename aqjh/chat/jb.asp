<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>竞标</title>
<script language="JavaScript">
function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}
function rc(list){if(list!="0"){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}}
</script>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>">
<center>
<p style="font-size:9pt; color:red">江湖商品竞标</p>
<a href="#" onClick="window.open('jbhelp.asp','jbsm','scrollbars=yes,resizable=no,width=320,height=280')" title="看看说明吧！"><b><font color="#FFFF00">帮助说明</font></b></a>
  <br>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM b where b='药品' or b='鲜花'  order  by -o",conn,2,2%>
  <table border="1" cellspacing="0" cellpadding="0" width="134" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC" align="center" height="8" >
    <tr bgcolor="#3399CC" onmouseout="this.bgColor='#3399CC';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td width=50%> 
        <div align="center">命令</div>
      </td>
      <td width=50%> 
        <div align="center">竞标</div>
      </td>
    </tr>
    <%do while not rs.eof and not rs.bof
if rs("q")=0 and rs("r")=true then   
	ti="无人竞标!"
elseif rs("n")<>"无" and rs("r")=False then
	ti="竞标人:"&rs("n")&"&#13&#10中标价:"&rs("q")&"&#13&#10现收入:"&rs("o")&"&#13&#10点击查看&#13&#10修改此物品"
else 
	ti="竞标人:"&rs("n")&"&#13&#10竟标价:"&rs("q")&"&#13&#10点击查看&#13&#10修改此物品"
end if
%>
    <tr bgcolor="#3399CC" onmouseout="this.bgColor='#3399CC';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td valign=top height="2"> 
        <div align="center"> <font color="#FF00FF"><a href="#" onClick="window.open('myjb.asp?wpname=<%=rs("a")%>','myjb','scrollbars=yes,resizable=no,width=420,height=330')" title="<%=ti%>"><%=rs("a")%></a></font></div>
      </td>
      <td valign=top height="2"> 
        <div align="center"><a href="javascript:s('/竞标$ <%=rs("a")%>|<%=(rs("q")+10000)%>')" title="<%=ti%>">竞标</a></div>
      </td>
    </tr>
    <%
rs.movenext
l=l+1
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
  </table>
</center>
</body>
</html>