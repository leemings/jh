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
<font color="0099FF">【发送告示】</font>
<hr noshade size="1" color=009900 width=50%>
<form method="post" action="gsok.asp" name="PostTopic">
 <div align="center">
<SELECT name=gslist onchange=DoTitle(this.options[this.selectedIndex].value)> 
<OPTION selected value="">快速指令</OPTION> 
<OPTION value="购买江湖会员，支持江湖发展！">购买江湖会员</OPTION>
<OPTION value="购买江湖程序、江湖空间，请登陆<a href=http://www.happyjh.com target=_blank>快乐官方网站</a>">购买江湖最新程序</OPTION>
<OPTION value="祝大家玩得开心^O^本站江湖为您提供良好的聊天环境！">祝大家玩得开心^O^</OPTION>
</SELECT><br>
    <br>
<input type="text" name="gs" size="50" maxlength="250">
<input type="submit" name="Submit" value="发送">
<input type="reset" name="reset" value="重写"></form></div>
</body>
</html>