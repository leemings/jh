<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
response.write "<HTML><HEAD><META http-equiv=Content-Type content='text/html; charset=gb2312'>"&_
		"<title>"& Forum_info(0) &"--表情转换列表</title>"&_
                "</HEAD><BODY "&Forum_body(11)&">"
%>
<!--#include file=inc/Forum_css.asp-->
<BR>
<div align="center">
<B><%=Forum_info(0)%>--≡ 心情转换 ≡</B></div>
<BR>
<table cellpadding=6 cellspacing=1 align=center class=tableborder1 >
    <tr >
      <th height="26" colspan="4" align="center">
         请点击图标添加到内容</A></th> 
    </tr>
    <tr >
      <td align="center" class=tablebody2 height="28" width="15%">代码</td>
      <td class=tablebody2 width="35%" height="28" align="center" >表情</td>
      <td class=tablebody2 align="center" height="28" width="15%">代码</td>
      <td class=tablebody2 width="35%" height="28" align="center" >表情</td>
    </tr><tr >
<%
dim ii,h
for i=0 to Forum_emotNum
h=h+1
	if len(i)=1 then ii="0" & i else ii=i
	response.write "<td class=tablebody1 align=center >[em"&ii&"]</td>"&_
			"<td class=tablebody1 align=center><img src="""&Forum_info(10)&Forum_emot(i)&""" border=0 onclick=""insertsmilie('[em"&ii&"]')"" style=""CURSOR: hand""></td>"
		if  h=2 then 
		response.write "</tr><tr>"
		h=0
		end if
next
%>
 <tr >
      <td colspan="4" class=tablebody2 >
        <p align="center"><input type=submit name="winclose" value="关 闭" onclick=window.close();></td>
    </tr>
  </table>
</body></html>
<script language="javascript">
function insertsmilie(smilieface){

	self.opener.frmAnnounce.Content.value+=smilieface;
}
</script>