<%@ LANGUAGE=VBScript codepage ="936" %>
<%
function bd(zy)
zy=replace(zy,chr(39),"")
zy=replace(zy,chr(34),"")
zy=replace(zy,",","��")
zy=Replace(zy,"<","")
zy=Replace(zy,">","")
zy=Replace(zy,"\x3c","")
zy=Replace(zy,"\x3e","")
zy=Replace(zy,"\074","")
zy=Replace(zy,"\74","")
zy=Replace(zy,"\75","")
zy=Replace(zy,"\76","")
zy=Replace(zy,"&lt","")
zy=Replace(zy,"&gt","")
zy=Replace(zy,"\076","")
badstr="�侫|��|ȥ��|��ʺ|����|����|����|��|����|������|����|��|ɵB|����|����|���|����|����|����|����|����|����|����|����|����|����|غ��|��ȥ|��Ů|����|����|���������|���������|���������|��Ƥ|��ͷ|��|�P|��|�H|����|��|��|����|����|����|����|����|����|ʮ����|��ү|������|���ϰ�|������|������|���|�Ҷ�|����|��|��|����|����|���׵���|һ��|�Ҹ�|�Ҳ�|��ĸ|����|��ð|��ï|�Ҹ�|ȥ��|����|���׹�|������|��ϯ|����"
bad=split(badstr,"|")
for i=0 to UBound(bad)
bd=Replace(zy,bad(i),"**")
next
end function
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"yamen")=0 then 
	Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');parent.history.go(-1);}</script>" 
	Response.End 
end if 

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
name=sjjh_name
diqu=bd(trim(request.form("diqu")))
nianling=bd(trim(request.form("nianling")))
email=bd(trim(request.form("email")))
oicq=bd(trim(request.form("oicq")))
touxiang=bd(trim(request.form("touxiang")))
qianming=bd(trim(request.form("qianming")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")
conn.Execute "update �û� set ����='"&diqu&"',����='"&nianling&"',����='"&email&"',Oicq='"&oicq&"',ͷ��='"&touxiang&"',ǩ��='"&qianming&"' where ����='"&Name&"'"
conn.close
set conn=nothing
%>
<html>
<title>���������޸�</title>
<body background="../JHIMG/BK_HC3W.GIF">
<table border=1 bgcolor="#BEE0FC" align=center width=336 cellpadding="10" cellspacing="13" height="293">
<tr valign="top">
<td bgcolor=#cccccc height="226">
<div align="center">
<p>&nbsp;</p>
<p><font color="#000000" size="3">���ɹ��޸�����!</font><font color="#FF3333" size="3"><br>
</font></p>
<p><font color="#FF3333" size="3"> </font> </p>
</div>
<table width=100%>
<tr>
<td>
<p align=center style='font-size:14;color:red'><font color="#000000"><br>
����޸��д�����������<br>
�ļ��ɣ�</font></p>
<p align=center>
<input type=button value=�رմ��� onClick='window.close()' name="button">
</p>
</td>
</tr>
</table>
</td>
</tr>
</table>
</html>
