<%@ LANGUAGE=VBScript.Encode codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if session("aqjh_name")=""  then Response.Redirect "../error.asp?id=440"
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%> - �������</title>
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
if(name=="" ){alert("��ʾ����������ֲ���Ϊ�գ�");return false;}
if(songname=="" ){alert("��ʾ����������Ϊ�գ�");return false;}
if(url=="" ){alert("��ʾ������·������Ϊ�գ�");return false;}
if(zhufu=="" ){alert("��ʾ��ף������Ϊ�գ�");return false;}
if(toname=="" ){alert("��ʾ���ո��˲���Ϊ�գ�");return false;}
}
</script>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<table align="center" width="98%">
          <form name="form" action="addsong.asp" method="post" onSubmit="return check(this);">
            <tr> 
              <td>�����: 
                <input class="smallInput" name="name" type="text" value="<%=aqjh_name%>" readonly></td>
            </tr>
            <tr> 
              <td>������: 
                <input class="smallInput" name="songname" type="text"></td>
            </tr>
            <tr> 
              <td>��·��: 
                <input class="smallInput" name="songurl" type="text" value="http://"></td>
            </tr>
            <tr> 
              <td>ף &nbsp���� 
                <input class="smallInput" name="zhufu" type="text"></td>
            </tr>
            <tr> 
              <td>���͸�: 
                <input class="smallInput" name="toname" type="text"></td>
            </tr>
            <tr> 
              <td><input class="buttonface" type="submit" name="Submit" value="�ύ"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                <input class="buttonface" type="reset" value="ȡ��"></td>
            </tr>
          </form>
        </table>
        </html>