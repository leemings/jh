	document.write("<span id='QQFace' style='position:absolute; width:0px; height:0px; z-index:99999; background-color: transparent; display: none; left: 0; top: 0;'></span> ");

	var sw	//���ݶ�ʱ
	var sWidth=300	//���������ʾ�Ŀ���
	var sHeight=300	//���������ʾ�ĸ߶�
	var sTime=6000	//���������ʾ��ʱ��

	//��ʾ����
	function ShowQQ(ID){
		clearTimeout(sw);

		document.getElementById("QQFace").style.width=sWidth
		document.getElementById("QQFace").style.height=sHeight

		document.getElementById("QQFace").style.left=document.body.scrollLeft+(document.body.clientWidth-sWidth)/2
		document.getElementById("QQFace").style.top=document.body.scrollTop+(document.body.clientHeight-sHeight)/2

		var QQStr="QQ/swf/" + ID + ".swf"
		document.getElementById("QQFace").innerHTML = '<OBJECT codeBase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" width="100%" height="100%"><PARAM NAME=movie VALUE="'+ QQStr +'"><param name=menu value=false><PARAM NAME=quality VALUE=high><PARAM NAME=play VALUE=false><param name="wmode" value="transparent"><embed src="' + QQStr +'" quality="high" wmode="transparent" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100%" height="100%"></embed>';
		QQFace.style.display=""

		sw=window.setTimeout("CloseQQ();",sTime);
	}

	//�رձ���
	function CloseQQ(){
		QQFace.style.display="none"
		clearTimeout(sw);
	}
//-->