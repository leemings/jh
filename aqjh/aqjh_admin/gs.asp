<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<SCRIPT LANGUAGE="JavaScript">
function DoTitle(addTitle) {
var revisedTitle;
var currentTitle = document.PostTopic.gs.value;
revisedTitle = addTitle;
document.PostTopic.gs.value=revisedTitle;
document.PostTopic.gs.focus();
return; }
</script>
<LINK href="css/css.css" type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<div align="center">
<font color="0099FF">�����͸�ʾ��</font>
<hr noshade size="1" color=009900 width=50%>
<form method="post" action="gsok.asp" name="PostTopic">
 <div align="center">
<SELECT name=gslist onchange=DoTitle(this.options[this.selectedIndex].value)> 
<OPTION selected value="">����ָ��</OPTION> 
<OPTION value="���򽭺���Ա��֧�ֽ�����չ��">���򽭺���Ա</OPTION>
<OPTION value="���򽭺����򡢽����ռ䣬���½<a href=http://www.7758530.com target=_blank>���ֹٷ���վ</a>">���򽭺����³���</OPTION>
<OPTION value="ף�����ÿ���^O^��վ����Ϊ���ṩ���õ����컷����">ף�����ÿ���^O^</OPTION>
</SELECT><br>
    <br>
<input type="text" name="gs" size="50" maxlength="250">
<input type="submit" name="Submit" value="����">
<input type="reset" name="reset" value="��д"></form></div>
</body>
</html>