<%@ LANGUAGE=VBScript.Encode codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if session("aqjh_name")=""  then Response.Redirect "../error.asp?id=440"
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%> - 自助点歌</title>
<LINK href=css.css type=text/css rel=stylesheet>
</head>
<SCRIPT language=JavaScript>
function DoTitle(addTitle) {
var revisedTitle;
var currentTitle = document.form.bds.value;
revisedTitle = addTitle;
document.form.bds.value=revisedTitle;
document.form.bds.focus();
return; }
function check(){
var name=document.form.name.value;
var songname=document.form.songname.value;
var songurl=document.form.songurl.value;
var zhufu=document.form.zhufu.value;
var toname=document.form.toname.value;
if(name=="" ){alert("提示：点歌人名字不能为空！");return false;}
if(songname=="" ){alert("提示：歌名不能为空！");return false;}
if(url=="" ){alert("提示：歌曲路径不能为空！");return false;}
if(zhufu=="" ){alert("提示：祝福不能为空！");return false;}
if(toname=="" ){alert("提示：收歌人不能为空！");return false;}
}
</script>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<table align="center" width="98%">
          <form name="form" action="addsong.asp" method="post" onSubmit="return check(this);">
            <tr> 
              <td>点歌人: 
                <input class="smallInput" name="name" type="text" value="<%=aqjh_name%>" readonly></td>
            </tr>
            <tr> 
              <td>歌曲名: 
                <input class="smallInput" name="songname" type="text"></td>
            </tr>
            <tr> 
              <td>歌路径: 
                <input class="smallInput" name="songurl" type="text" value="http://"></td>
            </tr>
            <tr> 
              <td>祝 &nbsp福： 
                <input class="smallInput" name="zhufu" type="text"></td>
            </tr>
            <tr> 
              <td>赠送给: 
                <input class="smallInput" name="toname" type="text"></td>
            </tr>
            <tr> 
              <td><input class="buttonface" type="submit" name="Submit" value="提交"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                <input class="buttonface" type="reset" value="取消"></td>
            </tr>
          </form>
        </table>
        </html>