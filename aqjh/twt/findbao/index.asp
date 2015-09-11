<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('刷钱是吗？喜欢是吗？点啊，刷啊！！');i=i+1;}top.location.href='../../exit.asp'}</script>
<HTML>
<HEAD>
<TITLE>无名孤岛</TITLE>
<META http-equiv=""Content-Type"" content=""text/html"">
<script language="JavaScript">
<!--
function MM_showHideLayers() { //v3.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
    obj.visibility=v; }
}

function MM_displayStatusMsg(msgStr) { //v1.0
  status=msgStr;
  document.MM_returnValue = true;
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
//-->
</script>
</HEAD>

<BODY BGCOLOR="#000000" leftmargin="0" topmargin="0">
<div id="Layer1" style="position:absolute; left:4px; top:73px; width:744px; height:475px; z-index:1"> 
  <form method=POST action='huayuan.asp' name=xiaohuayuan>
    <input name=h value=1 type=hidden>
    <div id="lt" style="position:absolute; left:243px; top:108px; width:72px; height:153px; z-index:1; visibility: hidden"> 
      <table width="100" border="1" cellspacing="0" cellpadding="0" bordercolor="#999999">
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=搜　索 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=休　息 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=嘻　戏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=破　坏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=迷　茫 type="submit" name=sh>
            </div>
          </td>
        </tr>
      </table>
    </div>
    <div id="ssx" style="position:absolute; left:400px; top:245px; width:72px; height:153px; z-index:1; visibility: hidden"> 
      <table width="100" border="1" cellspacing="0" cellpadding="0" bordercolor="#999999">
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=搜　索 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=休　息 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=嘻　戏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=破　坏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=迷　茫 type="submit" name=sh>
            </div>
          </td>
        </tr>
      </table>
    </div>
    <div id="xx" style="position:absolute; left:343px; top:113px; width:72px; height:153px; z-index:1; visibility: hidden"> 
      <table width="100" border="1" cellspacing="0" cellpadding="0" bordercolor="#999999">
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=搜　索 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=休　息 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=嘻　戏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=破　坏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=迷　茫 type="submit" name=sh>
            </div>
          </td>
        </tr>
      </table>
    </div>
    <div id="js" style="position:absolute; left:659px; top:210px; width:72px; height:153px; z-index:1; visibility: hidden"> 
      <table width="100" border="1" cellspacing="0" cellpadding="0" bordercolor="#999999">
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=搜　索 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=休　息 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=嘻　戏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=破　坏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> </div>
          </td>
        </tr>
      </table>
    </div>
    <div id="hysc" style="position:absolute; left:618px; top:254px; width:72px; height:153px; z-index:1; visibility: hidden"> 
      <table width="100" border="1" cellspacing="0" cellpadding="0" bordercolor="#999999">
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=搜　索 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=休　息 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=嘻　戏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=破　坏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=迷　茫 type="submit" name=sh>
            </div>
          </td>
        </tr>
      </table>
    </div>
    <div id="qcd" style="position:absolute; left:614px; top:45px; width:72px; height:153px; z-index:1; visibility: hidden"> 
      <table width="100" border="1" cellspacing="0" cellpadding="0" bordercolor="#999999">
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=搜　索 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=休　息 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=嘻　戏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=破　坏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=迷　茫 type="submit" name=sh>
            </div>
          </td>
        </tr>
      </table>
    </div>
    <div id="tud" style="position:absolute; left:204px; top:41px; width:72px; height:153px; z-index:1; visibility: hidden"> 
      <table width="100" border="1" cellspacing="0" cellpadding="0" bordercolor="#999999">
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=搜　索 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=休　息 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=嘻　戏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=破　坏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=迷　茫 type="submit" name=sh>
            </div>
          </td>
        </tr>
      </table>
    </div>
    <div id="du" style="position:absolute; left:124px; top:284px; width:72px; height:153px; z-index:1; visibility: hidden"> 
      <table width="100" border="1" cellspacing="0" cellpadding="0" bordercolor="#999999">
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=搜　索 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=休　息 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=嘻　戏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=破　坏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=迷　茫 type="submit" name=sh>
            </div>
          </td>
        </tr>
      </table>
    </div>
    <div id="han" style="position:absolute; left:580px; top:317px; width:72px; height:153px; z-index:1; visibility: hidden"> 
      <table width="100" border="1" cellspacing="0" cellpadding="0" bordercolor="#999999">
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=搜　索 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=休　息 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=嘻　戏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=破　坏 type="submit" name=sh>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFCC33"> 
          <td height="30"> 
            <div align="center" class="unnamed1"> 
              <input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=迷　茫 type="submit" name=sh>
            </div>
          </td>
        </tr>
      </table>
    </div>
    <table width="752" border="0" cellspacing="0" cellpadding="0" align="center" height="420">
      <tr> 
        <td width="46%"> 
          <table width="84%" border="0" cellspacing="0" cellpadding="0" align="center" height="356">
            <tr> 
              <td> 
                <div align="left" class="unnamed1"><font color="#FF0033">　</font></div>
              </td>
            </tr>
            <tr> 
              <td height="13">&nbsp; </td>
            </tr>
            <tr> 
              <td> 
                <div align="center">
                  <input onMouseOut="this.style.backgroundColor='#FFF7EE'" onMouseOver="this.style.backgroundColor='#FFCC33'" style="BACKGROUND-COLOR: #FFF7EE; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=   瀑布 type="text" name=sy readonly size=6 maxlength="8" onClick="MM_showHideLayers('lt','','hide','ssx','','hide','xx','','hide','js','','hide','hysc','','hide','qcd','','hide','du','','hide','tud','','show','han','','hide');MM_displayStatusMsg('ffff');return document.MM_returnValue">
                </div>
              </td>
            </tr>
            <tr> 
              <td height="51">&nbsp; </td>
            </tr>
            <tr> 
              <td> 
                <div align="center" class="unnamed1"> 
                  <input onMouseOut="this.style.backgroundColor='#FFF7EE'" onMouseOver="this.style.backgroundColor='#FFCC33'" style="BACKGROUND-COLOR: #FFF7EE; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value= 独龙潭 type="text" name=sy readonly size=8 maxlength="8"  onClick="MM_showHideLayers('lt','','show','ssx','','hide','xx','','hide','js','','hide','hysc','','hide','qcd','','hide','du','','hide','tud','','hide','han','','hide');MM_displayStatusMsg('ffff');return document.MM_returnValue">
                </div>
              </td>
            </tr>
            <tr> 
              <td>&nbsp; </td>
            </tr>
            <tr> 
              <td>&nbsp; </td>
            </tr>
            <tr> 
              <td>&nbsp; </td>
            </tr>
            <tr> 
              <td> 
                <div align="right">
                  <input onMouseOut="this.style.backgroundColor='#FFF7EE'" onMouseOver="this.style.backgroundColor='#FFCC33'" style="BACKGROUND-COLOR: #FFF7EE; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=山间小路 type="text" name=sy readonly size=10 maxlength="10" onClick="MM_showHideLayers('lt','','hide','ssx','','hide','xx','','hide','js','','hide','hysc','','hide','qcd','','hide','du','','show','tud','','hide','han','','hide');MM_displayStatusMsg('ffff');return document.MM_returnValue">
                </div>
              </td>
            </tr>
            <tr> 
              <td>&nbsp; </td>
            </tr>
          </table>
        </td>
        <td width="54%"> 
          <input onMouseOut="this.style.backgroundColor='#FFF7EE'" onMouseOver="this.style.backgroundColor='#FFCC33'" style="BACKGROUND-COLOR: #FFF7EE; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=  小河 type="text" name=sy readonly size=6 maxlength="8" onClick="MM_showHideLayers('lt','','hide','ssx','','hide','xx','','show','js','','hide','hysc','','hide','qcd','','hide','du','','hide','tud','','hide','han','','hide');MM_displayStatusMsg('ffff');return document.MM_returnValue">
          <table width="100%" border="0" cellspacing="0" cellpadding="0" height="356">
            <tr> 
              <td colspan="3"> 
                <div align="center"> 
                  <input onMouseOut="this.style.backgroundColor='#FFF7EE'" onMouseOver="this.style.backgroundColor='#FFCC33'" style="BACKGROUND-COLOR: #FFF7EE; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=    山之顶峰 type="text" name=sy readonly size=10 maxlength="14" onClick="MM_showHideLayers('lt','','hide','ssx','','hide','xx','','hide','js','','hide','hysc','','hide','qcd','','show','tud','','hide','du','','hide','han','','hide');MM_displayStatusMsg('ffff');return document.MM_returnValue">
                </div>
              </td>
            </tr>
            <tr> 
              <td colspan="2">&nbsp; </td>
              <td width="50%">&nbsp;</td>
            </tr>
            <tr> 
              <td colspan="3"> 
                <div align="center">　　　　　　 
                  <input onMouseOut="this.style.backgroundColor='#FFF7EE'" onMouseOver="this.style.backgroundColor='#FFCC33'" style="BACKGROUND-COLOR: #FFF7EE; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=山洞 type="text" name=sy readonly size=6 maxlength="8" onClick="MM_showHideLayers('lt','','hide','ssx','','hide','xx','','hide','js','','show','hysc','','hide','qcd','','hide','du','','hide','tud','','hide','han','','hide');MM_displayStatusMsg('ffff');return document.MM_returnValue">
                </div>
              </td>
            </tr>
            <tr> 
              <td width="31%">&nbsp; </td>
              <td width="36%"> 
                <input onMouseOut="this.style.backgroundColor='#FFF7EE'" onMouseOver="this.style.backgroundColor='#FFCC33'" style="BACKGROUND-COLOR: #FFF7EE; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value= 半山 type="text" name=sy readonly size=6 maxlength="8" onClick="MM_showHideLayers('lt','','hide','ssx','','hide','xx','','show','js','','hide','hysc','','hide','qcd','','hide','du','','hide','tud','','hide','han','','hide');MM_displayStatusMsg('ffff');return document.MM_returnValue">
              </td>
              <td width="33%">&nbsp;</td>
            </tr>
            <tr> 
              <td colspan="3">&nbsp; </td>
            </tr>
            <tr> 
              <td colspan="3"> 
                <div align="center"> 
                  <input onMouseOut="this.style.backgroundColor='#FFF7EE'" onMouseOver="this.style.backgroundColor='#FFCC33'" style="BACKGROUND-COLOR: #FFF7EE; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value= 山脚 type="text" name=sy readonly size=6 maxlength="8" onClick="MM_showHideLayers('lt','','hide','ssx','','show','xx','','hide','js','','hide','hysc','','hide','qcd','','hide','du','','hide','tud','','hide','han','','hide');MM_displayStatusMsg('ffff');return document.MM_returnValue">
                </div>
              </td>
            </tr>
            <tr> 
              <td colspan="3">&nbsp; </td>
            </tr>
            <tr> 
              <td colspan="3"> 
                <div align="center">
                  <input onMouseOut="this.style.backgroundColor='#FFF7EE'" onMouseOver="this.style.backgroundColor='#FFCC33'" style="BACKGROUND-COLOR: #FFF7EE; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=海岸 type="text" name=sy readonly size=6 maxlength="8" onClick="MM_showHideLayers('lt','','hide','ssx','','hide','xx','','hide','js','','hide','hysc','','hide','qcd','','hide','du','','hide','tud','','hide','han','','show');MM_displayStatusMsg('ffff');return document.MM_returnValue">
                </div>
              </td>
            </tr>
            <tr> 
              <td colspan="3">&nbsp; </td>
            </tr>
            <tr> 
              <td colspan="3">&nbsp; </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </form>
</div>
<div id="Layer2" style="position:absolute; left:47px; top:10px; width:692px; height:62px; z-index:2">
  <table width="715" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td> 
        <div align="center" class="unnamed1"><font size="2"><font color="#FFFF33"><%=aqjh_name%>坐船来到了无名的孤岛，只见四处荒无人烟，是祸是福也是未知之数，看来只有听天由命了！</font></div>
      </td>
    </tr>
  </table>
</div>
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=771 HEIGHT=600>
  <TR>
    <TD ROWSPAN=1 COLSPAN=1 WIDTH=260 HEIGHT=200><IMG SRC="images\gfdg行 1 x 列 1.jpg" WIDTH=260 HEIGHT=200 BORDER=0 ALT="gfdg行 1 x 列 1.jpg"></TD>
    <TD ROWSPAN=1 COLSPAN=1 WIDTH=260 HEIGHT=200><IMG SRC="images\gfdg行 1 x 列 2.jpg" WIDTH=260 HEIGHT=200 BORDER=0 ALT="gfdg行 1 x 列 2.jpg"></TD>
    <TD ROWSPAN=1 COLSPAN=1 WIDTH=251 HEIGHT=200><IMG SRC="images\gfdg行 1 x 列 3.jpg" WIDTH=258 HEIGHT=200 BORDER=0 ALT="gfdg行 1 x 列 3.jpg"></TD>
</TR>
<TR>
    <TD ROWSPAN=1 COLSPAN=1 WIDTH=260 HEIGHT=200><IMG SRC="images\gfdg行 2 x 列 1.jpg" WIDTH=260 HEIGHT=200 BORDER=0 ALT="gfdg行 2 x 列 1.jpg"></TD>
    <TD ROWSPAN=1 COLSPAN=1 WIDTH=260 HEIGHT=200><IMG SRC="images\gfdg行 2 x 列 2.jpg" WIDTH=260 HEIGHT=200 BORDER=0 ALT="gfdg行 2 x 列 2.jpg"></TD>
    <TD ROWSPAN=1 COLSPAN=1 WIDTH=251 HEIGHT=200><img src="images\gfdg行 2 x 列 3.jpg" width=258 height=200 border=0 alt="gfdg行 2 x 列 3.jpg"></TD>
</TR>
<TR>
    <TD ROWSPAN=1 COLSPAN=1 WIDTH=260 HEIGHT=200><IMG SRC="images\gfdg行 3 x 列 1.jpg" WIDTH=260 HEIGHT=200 BORDER=0 ALT="gfdg行 3 x 列 1.jpg"></TD>
    <TD ROWSPAN=1 COLSPAN=1 WIDTH=260 HEIGHT=200><IMG SRC="images\gfdg行 3 x 列 2.jpg" WIDTH=260 HEIGHT=200 BORDER=0 ALT="gfdg行 3 x 列 2.jpg"></TD>
    <TD ROWSPAN=1 COLSPAN=1 WIDTH=251 HEIGHT=200><IMG SRC="images\gfdg行 3 x 列 3.jpg" WIDTH=258 HEIGHT=200 BORDER=0 ALT="gfdg行 3 x 列 3.jpg"></TD>
</TR>
</TABLE>
</BODY>
</HTML>