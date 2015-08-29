<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
 if not master or session("flag")="" then
  Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
  call dvbbs_error()
 else
  dim body
  dim readme,Tlink
  call main()
  set rs=nothing
  conn.close
  set conn=nothing
 end if

 sub main()
%>
<TABLE > 
  <tr>
    <td height="30"> 
      <table cellpadding=3 cellspacing=1 border=0 width=100%>
        <tr >
        <td><font color=blue>欢迎<b><%=membername%></b>进入管理页面</font>
        </td>
        </tr>
            <tr >              
          <td width="100%" valign=top>
<%
dim objFSO
dim fdata
dim objCountFile
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if request("save")="" then
  Set objCountFile = objFSO.OpenTextFile(Server.MapPath("text.js"),1,True)
  If Not objCountFile.AtEndOfStream Then fdata = objCountFile.ReadAll
 else
  fdata=request("fdata")
  Set objCountFile=objFSO.CreateTextFile(Server.MapPath("text.js"),True)
  objCountFile.Write fdata
 end if
 objCountFile.Close
 Set objCountFile=Nothing
 Set objFSO = Nothing
%> 
<form method=post>
            <table width="100%" border="0" cellspacing="3" cellpadding="0">
              <tr> 
                <td width="3%" height="23"> </td>
                <td width="97%" height="23">论坛随机广告编辑：<br>
                      注意：文件将被保存在你安装目录下的<font color=red>test.js</font>文件里，如果您的空间不支持<font color=red>FSO</font>，请直接编辑该文件！<br>
                　　　如果您对html和js代码不了解，请只修改广告内容和超链接，其他不要随意编辑。</td>
              </tr>
              <tr> 
                <td width="3%">  </td>
                <td width="97%"> 
                  <textarea name="fdata" cols="110" rows="20"><%=fdata%></textarea>
                </td>
              </tr>
              <tr> 
                <td width="3%"> </td>
                <td width="97%">
                  <input type="reset" name="Reset" value="重新修改">
                  <input type="submit" name="save" value="提交修改"> 当前文件路径：<b><%=Server.MapPath("text.js")%></b>
                </td>
              </tr>
            </table></form>
            <p> </p>
            </td>
            </tr>
        </table>
        </td>
    </tr>
</table>
<%end sub%>