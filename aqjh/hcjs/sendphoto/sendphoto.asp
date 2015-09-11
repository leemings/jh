<%
if session("aqjh_name")=""  then
  Response.Redirect "../../error.asp?id=440"
else
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Namo WebEditor v4.0(Trial)">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>相片上传</title>
<Script language="javascript">
function mysubmit(theform)
{
    if(theform.big.value=="")
    {
    alert("请点击浏览按钮，选择您要上传的jpg或gif文件!")
    theform.big.focus;
    return (false);
    }
    else
    {
    str= theform.big.value;
    strs=str.toLowerCase();
    lens=strs.length;
    extname=strs.substring(lens-4,lens);
    if(extname!=".jpg" && extname!=".gif")
    {
    alert("请选择jpg或gif文件!");
    return (false);
    }
    }
    return (true);
}
</script>
<title>网友照片</title>
<meta name="generator" content="Microsoft FrontPage 4.0">
<link rel="stylesheet" href="../../css.CSS" type="text/css">
<SCRIPT language=JavaScript>
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  if (theURL != "fuckyou")
 {   window.open(theURL,winName,features);}
}
//-->
</SCRIPT></head>
<BODY background="../../bg.gif">
<div align=center> 
  <table cellspacing=0 cellpadding=0 width=505 border=0>
    <tbody><tr><td colspan=4 align=center>上传照片</td></tr> 
    <tr> 
      <td></td>
      <td width=483> 
        <table border="0" width="99%" height="182" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="65" valign="top"> 
              <form enctype="multipart/form-data" action="addpic.asp" method=post onSubmit="return mysubmit(this)">
                <table border="0" width="100%" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="277" bgcolor="#a5cf7b" height=24><font color="#FFFFFF"><b>网友相片上传</b></font></td>
                  </tr>
                  <tr> 
                    <td width="331"> <br><br>
                      <p align="center">大相片&nbsp;  
                        <input type="file" name="big" size="20">
                      </p>
                    </td>
                  </tr>
                  <tr> 
                    <td width="337"> <br>
                      <p align="center"> 
                        <input type="submit" value="上   传" name="B3">
                        　</p>
                    </td>
                  </tr>
                  <tr> 
                    <td><br><br>注意事项：<br>1、照片最佳尺寸为：400*300;200*150等，宽高比例为4：3时照片显示效果最佳。<br>2、严禁上传反动、黄色及一切违反国家法规的图片！一经发现，本站将梦少白删除其在江湖的帐号！<br>3、所有用户上传照片经站长审批后才能正常显示！
                    </td>
                  </tr>
                </table>
              </form>
              <br>
            </td>
          </tr>
        </table>
        <div align="center"><a href='photo.asp'>返回上页</a></div>
      </td>
      <td></td>
    </tr>
    <tr> 
      <td colspan="3"></td>
    </tr>
    </tbody> 
  </table>
</div>
</body>
<% end if %>
</html></script>