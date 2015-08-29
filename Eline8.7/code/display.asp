<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" type="text/css" href="cssasp.asp?id=3">
<style type="text/css">
<!--
.l           { background-color: #1892b5; height: 18px; border-left: 2px outset #66ccff; 
               border-right: 2px outset #66ccff; color:#000000;
               border-top: 2px outset #66ccff; padding-top: 2;height:18}
.lc           { background-color: #1892b5; height: 18px; border-left: 2px outset #66ccff; 
               border-right: 2px outset #66ccff; color:#000000;
               border-top: 2px outset #66ccff; padding-top: 2;height:20}
.l-h         { background-color: #1892b5; border-left: 2px outset #66ccff ; border-top: 2px outset #66ccff;color:#000000; }
.l-c         { background-color: #1892b5; border-top: 2px outset #66ccff }
.l-r         { background-color: #1892b5; border-right: 2px outset #66ccff; border-top: 2px outset #66ccff;color:#000000;}
.l-hc         { background-color: #1892b5; border-left: 2px outset #66ccff;color:#000000;}
.l-cc         { background-color: #1892b5;color:#000000; }
.l-rc         { background-color: #1892b5; border-right: 2px outset #66ccff;color:#000000;}
td{color:#000000;}
-->
</style>
<title>显示属性</title>

</head>

<body style="border:outset 1 #66ccff;margin: 0;background-color: #1892b5;padding:3px" scroll=no>
<SCRIPT LANGUAGE="JavaScript">
<!--
function restore()
{
td1.className="l";
td2.className="l";
td3.className="l";
td4.className="l";
td5.className="l";
td6.className="l";
td_1.className="l-h";
td_2.className="l-c";
td_3.className="l-c";
td_4.className="l-c";
td_5.className="l-c";
td_6.className="l-c";
style_window.style.display="none";
icon_window.style.display="none";
}
function c1()
{
td1.className="lc";
td_1.className="l-hc";
}
function c2()
{
td2.className="lc";
td_2.className="l-cc";
}
function c3()
{
td3.className="lc";
td_3.className="l-cc";
style_window.style.display="block";
}
function c4()
{
td4.className="lc";
td_4.className="l-cc";
}
function c5()
{
td5.className="lc";
td_5.className="l-cc";
icon_window.style.display="block";
}
function c6()
{
td6.className="lc";
td_6.className="l-cc";
}
//-->
</SCRIPT>
<form name="free" method="post" target="_target">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="400" height="61">
    <tr>
      <td width="50" height="20" align="center" valign="bottom" onclick=restore();c1();><div id="td1" class="lc">背景</div></td>
      <td width="50" height="20" align="center" valign="bottom" onclick=restore();c2();><div id="td2" class="l">屏保</div></td>
      <td width="50" height="20" align="center" valign="bottom" onclick=restore();c3();><div id="td3" class="l">外观</div></td>
      <td width="50" height="20" align="center" valign="bottom" onclick=restore();c4();><div id="td4" class="l">Web</div></td>
      <td width="50" height="20" align="center" valign="bottom" onclick=restore();c5();><div id="td5" class="l">效果</div></td> 
      <td width="50" height="20" align="center" valign="bottom" onclick=restore();c6();><div id="td6" class="l">设置</div></td>
      <td width="50" height="20" align="center" valign="bottom" ></td>
      <td width="50" height="20" align="center" valign="bottom"></td>
    </tr>
    <tr  style="">
      <td width="50" height="1" align="center" class="l-hc" id="td_1" >&nbsp;</td>
      <td width="50" height="1" align="center" class="l-c" id="td_2">&nbsp;</td>
      <td width="50" height="1" align="center" class="l-c" id="td_3">&nbsp;</td>
      <td width="50" height="1" align="center" class="l-c" id="td_4">&nbsp;</td>
      <td width="50" height="1" align="center" class="l-c" id="td_5">&nbsp;</td> 
      <td width="50" height="1" align="center" class="l-c" id="td_6">&nbsp;</td>
      <td width="50" height="1" align="center" class="l-c" id="td_">&nbsp;</td>
      <td width="50" height="1" align="center" class="l-r" id="td_">&nbsp;</td>
    </tr>
    <tr>
      <td width="398" height="360" class="up" colspan="8" style="border-top-style: solid; border-top-width: 0; padding: 5">
        <p align="center">&nbsp;
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" height="325">
            <tr>
              <td width="100%" height="280">
             
  <div align="center">
             
  <table border="0" cellpadding="0" cellspacing="1" width="365" height="200"  id="style_window" style="display:none">
    <tr>
      <td width="363" height="136" colspan="2" class="down1">
<Iframe src="styleshow.asp" name="style_show" frameborder=0 scrolling=no  height="158" width="363"></Iframe></td>
    </tr>
    <tr>
      <td width="180" height="14">方案(S)</td>
      <td width="185" height="14"></td>
    </tr>
    <tr>
      <td width="180" height="14">
	  
	  <select size="1" name="cssname" style="width:100%" onchange=form._apply.disabled=false;parent.style_show.location="styleshow.asp?id="+cssname.options[cssname.selectedIndex].value>
		
<option value="1" >Windows 默认</option>

<option value="2" >Windows 褐色</option>

<option value="3" >Windows 蓝色</option>

<option value="4" >Windows xp</option>

<option value="5" >Windows 黄色</option>

<option value="6" >Windows 绿色</option>

<option value="7" >Windows 黑色</option>

<option value="8" >Windows 土黄</option>

        </select></td>
      <td width="185" height="14"></td>
    </tr>
    <tr>
      <td width="180" height="14">项目(I):</td>
      <td width="185" height="14">大小(Z): 颜色(L)</td>    
    </tr>
    <tr>
      <td width="180" height="14"><select size="1" name="D1" style="width:100%">
        <option value="1">桌面</option>
        &nbsp;    
        </select></td>
      <td width="185" height="14"></td>
    </tr>
    <tr>
      <td width="180" height="14">字体(F):</td>
      <td width="185" height="14">大小(Z): 颜色(L)</td>    
    </tr>
    <tr>
      <td width="180" height="14"><select size="1" name="D1" style="width:100%">
        <option value="1">宋体</option>
        &nbsp;    
        </select></td>
      <td width="185" height="14"><select size="1" name="D1">
        <option value="1">宋体</option>
        &nbsp;    
        </select><select size="1" name="D1">
        <option value="1">宋体</option>
        &nbsp;    
        </select></td>
    </tr>
  </table>
  
<table border="0" cellpadding="0" cellspacing="1" width="365" height="226" id="icon_window" style="display: none">
  <tr>
    <td width="363" height="94" colspan="2" class="down1"><Iframe src="iconshow.asp" name="icon_show" frameborder=0 scrolling=no  height="100%" width="363">
      </Iframe>
    </td>
  </tr>
  <tr>
    <td width="67" height="86">图标(I)</td>
    <td width="298" height="86">
	  
	<select size="1" name="iconname" style="width:100%" onchange="form._apply.disabled=false;parent.icon_show.location=&quot;iconshow.asp?id=&quot;+iconname.options[iconname.selectedIndex].id">
        
        <option id="2" value="Apples" >Apples</option>
        
      </select></td>
  </tr>
</table>

  </div>
  <a id="hw_url" href="index.htm" target="_desktop" sytle="display:none"></a>
             </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
  </center>
    <tr>
      <td width="398" height="37" colspan="8" style="border-top-style: solid; border-top-width: 0; padding: 5">
        <p align="right"><input type="button" value="确定" name="_ok" style="width: 60;height:22" class=up onclick="parent._target.location='displaysave.asp?css='+cssname.options[cssname.selectedIndex].value+'&icon='+iconname.options[iconname.selectedIndex].value;form._apply.disabled=true;alert('请刷新桌面');window.close();" onhelp=alert("保存所有更改，关闭对话框。")> 
        <input type="button" value="取消" name="_cancel" style="width: 60;height:22" class=up onclick="window.close();" onhelp=alert("关闭对话框，不保存所有更改。")> 
        <input type="button" value="应用(A)" name="_apply" style="width: 60;height:22" disabled class=up onclick="parent._target.location='displaysave.asp?css='+cssname.options[cssname.selectedIndex].value+'&icon='+iconname.options[iconname.selectedIndex].value;form._apply.disabled=true;alert('请刷新桌面');" onhelp=alert("保存所有更改，但不关闭对话框。")>
      </td>
    </tr>
  </table>
</div>
</form>
<iframe name=_target style='display:none'></iframe> 
</body>

</html>
