<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
id=request.querystring("id")
tianqi=request.form("tianqi")
xinqing=request.form("xinqing")
title=request.form("title")
diary=request.form("diary")
tianqi=LCase(trim(Request.form("tianqi")))
tianqi=replace(tianqi,"'","")
tianqi=replace(tianqi,chr(34),"")
tianqi=Replace(tianqi,"<","")
tianqi=Replace(tianqi,">","")
tianqi=Replace(tianqi,"\x3c","")
tianqi=Replace(tianqi,"\x3e","")
tianqi=Replace(tianqi,"\074","")
tianqi=Replace(tianqi,"\74","")
tianqi=Replace(tianqi,"\75","")
tianqi=Replace(tianqi,"\76","")
tianqi=Replace(tianqi,"&lt","")
tianqi=Replace(tianqi,"&gt","")
tianqi=Replace(tianqi,"\076","")
badstr="�侫|��|ȥ��|��ʺ|����|����|����|��|����|������|����|��|ɵB|����|����|���|����|����|����|����|����|����|����|����|����|����|غ��|��ȥ |���������|���������|���������|��Ƥ|��ͷ|��|�P|��|�H|����|��|��|����|����|����|����|����|��Ů|����|ʮ����|��ү|���|�Ҷ�|����|��|��|asp|com|net|www|xajh|202|61|jh|����|or|261|����|����"
bad=split(badstr,"|")
for i=0 to UBound(bad)
tianqi=Replace(tianqi,bad(i),"**")
next
xinqing=LCase(trim(Request.form("xinqing")))
xinqing=replace(xinqing,"'","")
xinqing=replace(xinqing,chr(34),"")
xinqing=Replace(xinqing,"<","")
xinqing=Replace(xinqing,">","")
xinqing=Replace(xinqing,"\x3c","")
xinqing=Replace(xinqing,"\x3e","")
xinqing=Replace(xinqing,"\074","")
xinqing=Replace(xinqing,"\74","")
xinqing=Replace(xinqing,"\75","")
xinqing=Replace(xinqing,"\76","")
xinqing=Replace(xinqing,"&lt","")
xinqing=Replace(xinqing,"&gt","")
xinqing=Replace(xinqing,"\076","")
badstr="�侫|��|ȥ��|��ʺ|����|����|����|��|����|������|����|��|ɵB|����|����|���|����|����|����|����|����|����|����|����|����|����|غ��|��ȥ |���������|���������|���������|��Ƥ|��ͷ|��|�P|��|�H|����|��|��|����|����|����|����|����|��Ů|����|ʮ����|��ү|���|�Ҷ�|����|��|��|asp|com|net|www|xajh|202|61|jh|����|or|261|����|����"
bad=split(badstr,"|")
for i=0 to UBound(bad)
xinqing=Replace(xinqing,bad(i),"**")
next
title=LCase(trim(Request.form("title")))
title=replace(title,"'","")
title=replace(title,chr(34),"")
title=Replace(title,"<","")
title=Replace(title,">","")
title=Replace(title,"\x3c","")
title=Replace(title,"\x3e","")
title=Replace(title,"\074","")
title=Replace(title,"\74","")
title=Replace(title,"\75","")
title=Replace(title,"\76","")
title=Replace(title,"&lt","")
title=Replace(title,"&gt","")
title=Replace(title,"\076","")
badstr="�侫|��|ȥ��|��ʺ|����|����|����|��|����|������|����|��|ɵB|����|����|���|����|����|����|����|����|����|����|����|����|����|غ��|��ȥ |���������|���������|���������|��Ƥ|��ͷ|��|�P|��|�H|����|��|��|����|����|����|����|����|��Ů|����|ʮ����|��ү|���|�Ҷ�|����|��|��|asp|com|net|www|xajh|202|61|jh|����|or|261|����|����"
bad=split(badstr,"|")
for i=0 to UBound(bad)
title=Replace(title,bad(i),"**")
next
diary=LCase(trim(Request.form("diary")))
diary=replace(diary,"'","")
diary=replace(diary,chr(34),"")
diary=Replace(diary,"<","")
diary=Replace(diary,">","")
diary=Replace(diary,"\x3c","")
diary=Replace(diary,"\x3e","")
diary=Replace(diary,"\074","")
diary=Replace(diary,"\74","")
diary=Replace(diary,"\75","")
diary=Replace(diary,"\76","")
diary=Replace(diary,"&lt","")
diary=Replace(diary,"&gt","")
diary=Replace(diary,"\076","")
badstr="�侫|��|ȥ��|��ʺ|����|����|����|��|����|������|����|��|ɵB|����|����|���|����|����|����|����|����|����|����|����|����|����|غ��|��ȥ |���������|���������|���������|��Ƥ|��ͷ|��|�P|��|�H|����|��|��|����|����|����|����|����|��Ů|����|ʮ����|��ү|���|�Ҷ�|����|��|��|asp|com|net|www|xajh|202|61|jh|����|or|261|����|����"
bad=split(badstr,"|")
for i=0 to UBound(bad)
diary=Replace(diary,bad(i),"**")
next
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")
conn.Execute "update j set c='"&xinqing&"',d='"&tianqi&"',e='"&title&"',f='"&diary&"' where id="&id
conn.close
set conn=nothing
%>
<html>
<title>�ռ���д�ɹ�����wWw.happyjh.com��</title>
<body background="../bgcheetah.gif">
<table border=1 bgcolor="#BEE0FC" align=center width=336 cellpadding="10" cellspacing="13" height="293">
<tr valign="top">
<td background="../bgcheetah.gif" height="226">
<div align="center">
<p>&nbsp;</p>
<p><font color="#000000" size="3">�ռ��޸ĳɹ���</font><font color="#FF3333" size="3"><br>
</font></p>
<p><font color="#FF3333" size="3"> </font> </p>
</div>
<table width=100%>
<tr>
<td>
<p align=center style='font-size:14;color:red'><font color="#000000"><br>
�������ռǴ����Բ鿴�Լ�д��<br>
</font><font color="#000000">�ռǣ�<br>
</font></p>
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