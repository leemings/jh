<HTML>
  <HEAD>
    <TITLE>在线传讯</TITLE>
<!--
.link1 {  color: #FFFFFF}
.link2 {  color: #FF0000}
.link3 {  color: #006600}
.big1 {  font-size: 14px; text-align: justify; vertical-align: 1%; line-height: 18px}
.title {  font-size: 14px; background-color: #CCFFCC; font-weight: bold}
.botton {  color: #000066; background-color: #FFCC99; font-size: 12px; cursor: hand;}
-->
</style>
<link rel="stylesheet" href="../3508.css">

  </HEAD>
  <SCRIPT LANGUAGE="VBScript">
    Sub OpenChat()
      Window.Open "Calltwo.asp", "ChaWindow", "MENUBAR=0, SCROLLBARS=0, WIDTH=400, HEIGHT=150"
      Window.Close
    End Sub
  </SCRIPT>
  
<BODY bgcolor="#ffffdf">
<div align="left"></div>
<table width="100%" border="0" cellspacing="0" align="center">
  <tr>
    <td> 
      <div align="center"><img src="../image/ad.gif" width="390" height="30"></div>
    </td>
  </tr>
  <tr>
    <td> 
      <div align="left"></div>
      <table width="100%" bgcolor="#acc4b0" bordercolorlight="#acc4b0" cellpadding="0" border="0">
        <tr>
          <td height="73">
            <table width="100%" border="0" cellspacing="0" height="100%">
              <tr bgcolor="#FFFFFF"> 
                <td><%= Application("ShowMsg") %></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <form>
        <div align="right">
          <input   class="botton"   type="BUTTON" value="关闭窗口" onClick="VBScript:Window.Close" name="BUTTON">
          <input   class="botton"   type="BUTTON" value="回复信息" onClick="OpenChat" name="BUTTON2">
        </div>
      </form>
      <% Application("ShowMsg") = Empty %> </td>
  </tr>
</table>
</BODY>
</HTML>








