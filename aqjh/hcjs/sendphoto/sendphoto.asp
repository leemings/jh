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
<title>��Ƭ�ϴ�</title>
<Script language="javascript">
function mysubmit(theform)
{
    if(theform.big.value=="")
    {
    alert("���������ť��ѡ����Ҫ�ϴ���jpg��gif�ļ�!")
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
    alert("��ѡ��jpg��gif�ļ�!");
    return (false);
    }
    }
    return (true);
}
</script>
<title>������Ƭ</title>
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
    <tbody><tr><td colspan=4 align=center>�ϴ���Ƭ</td></tr> 
    <tr> 
      <td></td>
      <td width=483> 
        <table border="0" width="99%" height="182" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="65" valign="top"> 
              <form enctype="multipart/form-data" action="addpic.asp" method=post onSubmit="return mysubmit(this)">
                <table border="0" width="100%" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="277" bgcolor="#a5cf7b" height=24><font color="#FFFFFF"><b>������Ƭ�ϴ�</b></font></td>
                  </tr>
                  <tr> 
                    <td width="331"> <br><br>
                      <p align="center">����Ƭ&nbsp;  
                        <input type="file" name="big" size="20">
                      </p>
                    </td>
                  </tr>
                  <tr> 
                    <td width="337"> <br>
                      <p align="center"> 
                        <input type="submit" value="��   ��" name="B3">
                        ��</p>
                    </td>
                  </tr>
                  <tr> 
                    <td><br><br>ע�����<br>1����Ƭ��ѳߴ�Ϊ��400*300;200*150�ȣ���߱���Ϊ4��3ʱ��Ƭ��ʾЧ����ѡ�<br>2���Ͻ��ϴ���������ɫ��һ��Υ�����ҷ����ͼƬ��һ�����֣���վ�����ٰ�ɾ�����ڽ������ʺţ�<br>3�������û��ϴ���Ƭ��վ�����������������ʾ��
                    </td>
                  </tr>
                </table>
              </form>
              <br>
            </td>
          </tr>
        </table>
        <div align="center"><a href='photo.asp'>������ҳ</a></div>
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