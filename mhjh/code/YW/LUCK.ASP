<%
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
%>
<html>
<head>
<title>爱情速配中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">
<!--
-->
</style>
<link rel="stylesheet" href="../../style.css">

<script language="JavaScript">
<!--//

//设置下面一些参数，小球移动速度1-50，数值大速度快；
var ballWidth = 98;
var ballHeight = 80;
var BallSpeed = 30;

var maxBallSpeed = 50;
var xMax;
var yMax;
var xPos = 0;
var yPos = 0;
var xDir = 'right';  //水平方向向右移动
var yDir = 'down'; //垂直方向向下移动
var superballRunning = true;
var tempBallSpeed;
var currentBallSrc;
var newXDir;
var newYDir;

function initializeBall() {
   if (document.all) {
      xMax = document.body.clientWidth
      yMax = document.body.clientHeight
      document.all("superball").style.visibility = "visible";
      }
   else if (document.layers) {
      xMax = window.innerWidth;
      yMax = window.innerHeight;
      document.layers["superball"].visibility = "show";
      }
   setTimeout('moveBall()',400);
   }

function moveBall() {
   if (superballRunning == true) {
      calculatePosition();
      if (document.all) {
         document.all("superball").style.left = xPos + document.body.scrollLeft;
         document.all("superball").style.top = yPos + document.body.scrollTop;
         }
      else if (document.layers) {
         document.layers["superball"].left = xPos + pageXOffset;
         document.layers["superball"].top = yPos + pageYOffset;
         }
      setTimeout('moveBall()',30);
      }
   }

function calculatePosition() {
   if (xDir == "right") {
      if (xPos > (xMax - ballWidth - BallSpeed)) { 
         xDir = "left";
         }
      }
   else if (xDir == "left") {
      if (xPos < (0 + BallSpeed)) {
         xDir = "right";
         }
      }
   if (yDir == "down") {
      if (yPos > (yMax - ballHeight - BallSpeed)) {
         yDir = "up";
         }
      }
   else if (yDir == "up") {
      if (yPos < (0 + BallSpeed)) {
         yDir = "down";
         }
      }
   if (xDir == "right") {
      xPos = xPos + BallSpeed;
      }
   else if (xDir == "left") {
      xPos = xPos - BallSpeed;
      }
   else {
      xPos = xPos;
      }
   if (yDir == "down") {
      yPos = yPos + BallSpeed;
      }
   else if (yDir == "up") {
      yPos = yPos - BallSpeed;
      }
   else {
      yPos = yPos;
      }
   }

if (document.all||document.layers)
window.onload = initializeBall;
window.onresize = new Function("window.location.reload()");

// -->
</script>
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body background="riji/014.gif">
<p align="center">
<img border="0" src="riji/xiaoxiong.gif"><img border="0" src="riji/beibei.gif"><span id="superball"><img name="superballImage" src="../image/trailgif1.gif" height="30" width="30" border="0" "></span>
      
      <p align="center">[爱情速配中心公告]<br>
        <br>
        本中心对速配成功的顾客收取5000两，速配成功后还将免费为你刊登求婚信息。让有情人终成眷淑。<br>
        <br>
        <input onClick="javascript:window.document.location.href='luck1.asp'" title=开始测算 type=button value="开始速配" name="button7">
      </p>
<p align="center">　</p>
<p align="center">　</p>
</body>

</html>


