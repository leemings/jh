<script language="JavaScript">
	var appFlag=true;
	if(navigator.appName != "Microsoft Internet Explorer") {
		appFlag=false;
	} else {
		var pos1=navigator.appVersion.indexOf("MSIE");
		if(pos1<0) {
			appFlag=false;
		} else {
			var pos2=navigator.appVersion.substring(pos1,navigator.appVersion.length-1).indexOf(";");
			if(pos2<0) {
				appFlag=false;
			} else {
				var appVer=navigator.appVersion.substring(pos1+4,pos1+pos2);
				if(appVer<5.5) appFlag=false;
			}
		}
	}
	if(!appFlag) alert("建议使用IE5.5以上版本的浏览器，否则无法正常使用照相馆");
	var photoObj;
	var userCount=1;
	var MyBg="<%=GetBackground(request.cookies("myshow_"&userid))%>";
	var deg2radians=Math.PI*2/360;
	
	function process() {
		if(event.keyCode==40||event.keyCode==37||event.keyCode==39||event.keyCode==38) {
			var deltaX=0,deltaY=0;
			if(!event.ctrlKey) {
				if(event.keyCode==37) deltaX=-1;
				if(event.keyCode==38) deltaY=-1;
				if(event.keyCode==39) deltaX=1;
				if(event.keyCode==40) deltaY=1;
			} else {
				if(event.keyCode==37) deltaX=-10;
				if(event.keyCode==38) deltaY=-10;
				if(event.keyCode==39) deltaX=10;
				if(event.keyCode==40) deltaY=10;
			}
			if(document.all.photo_mode[0].checked) {
				var TempWidth,TempHeight;
				if(document.all.rot_deg[0].checked || document.all.rot_deg[2].checked) {
					TempWidth=photoObj.style.pixelWidth;
					if(photoObj.id.substr(0,4)=="Body"||photoObj.id.substr(0,12)=="Accouterment") {
						TempHeight=photoObj.style.pixelHeight;
					} else {
						TempHeight=20;
					}
				} else {
					TempHeight=photoObj.style.pixelWidth;
					if(photoObj.id.substr(0,4)=="Body"||photoObj.id.substr(0,12)=="Accouterment") {
						TempWidth=photoObj.style.pixelHeight;
					} else {
						TempWidth=20;
					}
				}
				if(photoObj.style.pixelLeft+deltaX>=20-TempWidth&&photoObj.style.pixelLeft+deltaX<=document.all.ThePhoto.style.pixelWidth-22) photoObj.style.pixelLeft=photoObj.style.pixelLeft+deltaX;
				if(photoObj.style.pixelTop+deltaY>=20-TempHeight&&photoObj.style.pixelTop+deltaY<=document.all.ThePhoto.style.pixelHeight-22) photoObj.style.pixelTop=photoObj.style.pixelTop+deltaY;
			} else
				if(document.all.photo_mode[1].checked) {
					if(photoObj.style.pixelWidth-deltaX>=20 &&
						photoObj.children[0].style.pixelLeft-deltaX<=0 &&
						photoObj.style.pixelLeft+deltaX<=document.all.ThePhoto.style.pixelWidth-22) {
						photoObj.children[0].style.pixelLeft=photoObj.children[0].style.pixelLeft-deltaX;
						photoObj.style.pixelWidth=photoObj.style.pixelWidth-deltaX;
						photoObj.style.pixelLeft=photoObj.style.pixelLeft+deltaX;
					}
					if(photoObj.style.pixelHeight-deltaY>=20 &&
						photoObj.children[0].style.pixelTop-deltaY<=0 &&
						photoObj.style.pixelTop+deltaY<=document.all.ThePhoto.style.pixelHeight-22) {
						photoObj.children[0].style.pixelTop=photoObj.children[0].style.pixelTop-deltaY;
						photoObj.style.pixelHeight=photoObj.style.pixelHeight-deltaY;
						photoObj.style.pixelTop=photoObj.style.pixelTop+deltaY;
					}
				} else
				if(document.all.photo_mode[2].checked) {
					if(photoObj.style.pixelWidth+deltaX>=20-photoObj.style.pixelLeft &&
						photoObj.style.pixelWidth+deltaX>=20 &&
						photoObj.style.pixelWidth+deltaX<=photoObj.children[0].style.pixelLeft+photoObj.children[0].style.pixelWidth-1) {
						photoObj.style.pixelWidth=photoObj.style.pixelWidth+deltaX;
					}
					if(photoObj.style.pixelHeight+deltaY>=20-photoObj.style.pixelTop &&
						photoObj.style.pixelHeight+deltaY>=20 &&
						photoObj.style.pixelHeight-1+deltaY<=photoObj.children[0].style.pixelTop+photoObj.children[0].style.pixelHeight-1) {
						photoObj.style.pixelHeight=photoObj.style.pixelHeight+deltaY;
					}
				} else
					if(document.all.photo_mode[3].checked) {
						var OldValue;
						if(photoObj.children[0].style.pixelWidth+deltaX>=70 &&
							photoObj.children[0].style.pixelWidth+deltaX<=280) {
							OldValue=photoObj.children[0].style.pixelWidth;
							photoObj.children[0].style.pixelWidth=photoObj.children[0].style.pixelWidth+deltaX;
							photoObj.children[0].style.pixelLeft=Math.round(photoObj.children[0].style.pixelLeft*photoObj.children[0].style.pixelWidth/OldValue);
							for(var i=0;i<photoObj.children[0].children.length;i++) photoObj.children[0].children[i].style.pixelWidth=photoObj.children[0].children[i].style.pixelWidth+deltaX;
							photoObj.style.pixelWidth=Math.round(photoObj.style.pixelWidth*photoObj.children[0].style.pixelWidth/OldValue);
						}
						if(photoObj.children[0].style.pixelHeight+deltaY>=113 &&
							photoObj.children[0].style.pixelHeight+deltaY<=452) {
							OldValue=photoObj.children[0].style.pixelHeight;
							photoObj.children[0].style.pixelHeight=photoObj.children[0].style.pixelHeight+deltaY;
							photoObj.children[0].style.pixelTop=Math.round(photoObj.children[0].style.pixelTop*photoObj.children[0].style.pixelHeight/OldValue);
							for(var i=0;i<photoObj.children[0].children.length;i++) photoObj.children[0].children[i].style.pixelHeight=photoObj.children[0].children[i].style.pixelHeight+deltaY;
							photoObj.style.pixelHeight=Math.round(photoObj.style.pixelHeight*photoObj.children[0].style.pixelHeight/OldValue);
						}
					}
			return false;
		}
	}
	
	function forward() {
		var OldzIndex=photoObj.style.zIndex;
		var CurzIndex;
		for(var i=2;i<document.all.ThePhoto.children.length;i++) {
			if(document.all.ThePhoto.children[i].id.substr(0,11)=="BodyBorder_") {
				if(document.all.ThePhoto.children[i].style.zIndex==OldzIndex+20) {
					CurzIndex=OldzIndex+20;
					for(var j=0;j<document.all.ThePhoto.children[i].children[0].children.length;j++) 
						document.all.ThePhoto.children[i].children[0].children[j].style.zIndex=OldzIndex+2+j;
					document.all.ThePhoto.children[i].children[0].style.zIndex=OldzIndex+1;
					document.all.ThePhoto.children[i].style.zIndex=OldzIndex;
					for(var j=0;j<photoObj.children[0].children.length;j++) 
						photoObj.children[0].children[j].style.zIndex=CurzIndex+2+j;
					photoObj.children[0].style.zIndex=CurzIndex+1;
					photoObj.style.zIndex=CurzIndex;
					break;
				}
			}
		}
	}

	function reverse() {
		var OldzIndex=photoObj.style.zIndex;
		var CurzIndex;
		for(var i=2;i<document.all.ThePhoto.children.length;i++) {
			if(document.all.ThePhoto.children[i].id.substr(0,11)=="BodyBorder_") {
				if(document.all.ThePhoto.children[i].style.zIndex==OldzIndex-20) {
					CurzIndex=OldzIndex-20;
					for(var j=0;j<document.all.ThePhoto.children[i].children[0].children.length;j++) 
						document.all.ThePhoto.children[i].children[0].children[j].style.zIndex=OldzIndex+2+j;
					document.all.ThePhoto.children[i].children[0].style.zIndex=OldzIndex+1;
					document.all.ThePhoto.children[i].style.zIndex=OldzIndex;
					for(var j=0;j<photoObj.children[0].children.length;j++) 
						photoObj.children[0].children[j].style.zIndex=CurzIndex+2+j;
					photoObj.children[0].style.zIndex=CurzIndex+1;
					photoObj.style.zIndex=CurzIndex;
					break;
				}
			}
		}
	}

	function selectBody(selvalue) {
  document.all.SelName.selectedIndex=0;
  selName("");
  if(userCount>1){
   document.all.SelAccouterment.selectedIndex=0;
   selAccouterment("");
  }
  selBody(selvalue);
 }
	
	function selBody(selValue) {
		photoObj=document.all.ThePhoto;
		for(var i=0;i<photoObj.children.length;i++) 
			photoObj.children[i].style.borderStyle="";
		if(selValue!="") {
			photoObj=eval("document.all.BodyBorder_"+selValue);
			photoObj.style.borderWidth="1px";
			photoObj.style.borderStyle="dotted";
			photoObj.style.borderColor="#ff0000";
			document.onkeydown=process;
			document.all.photo_mode[0].disabled=false;
			document.all.photo_mode[1].disabled=false;
			document.all.photo_mode[2].disabled=false;
			document.all.photo_mode[3].disabled=false;
			document.all.photo_front.disabled=false;
			document.all.photo_rear.disabled=false;
			document.all.mirror.disabled=false;
			if(photoObj.filters.item(0).Mirror!=0) {
				document.all.mirror.checked=true;
			} else {
				document.all.mirror.checked=false;
			}
		} else {
			document.onkeydown=new Function("");
			document.all.photo_mode[0].disabled=true;
			document.all.photo_mode[0].checked=true;
			document.all.photo_mode[1].disabled=true;
			document.all.photo_mode[2].disabled=true;
			document.all.photo_mode[3].disabled=true;
			document.all.photo_front.disabled=true;
			document.all.photo_rear.disabled=true;
			document.all.mirror.disabled=true;
			document.all.mirror.checked=false;
		}
	}
	
	function selectName(selvalue) {
  document.all.SelBody.selectedIndex=0;
  selBody("");
  if(userCount>1){
   document.all.SelAccouterment.selectedIndex=0;
   selAccouterment("");
  }
  selName(selvalue);
 }

	function selName(selValue) {
		photoObj=document.all.ThePhoto;
		for(var i=0;i<photoObj.children.length;i++) 
			photoObj.children[i].style.borderStyle="";
		if(selValue!="") {
			if(selValue=="-1") {
				photoObj=document.all.PhotoNameBorder;
			} else {
				if(selValue=="-2") {
					photoObj=document.all.DateBorder;
				} else {
					photoObj=eval("document.all.NameBorder_"+selValue);
				}
			}
			photoObj.style.borderWidth="1px";
			photoObj.style.borderStyle="dotted";
			photoObj.style.borderColor="#0000ff";
			document.onkeydown=process;
			document.all.photo_mode[0].disabled=false;
			document.all.photo_mode[1].disabled=true;
			document.all.photo_mode[2].disabled=true;
			document.all.photo_mode[3].disabled=true;
			document.all.selFont.disabled=false;
			for(var i=0;i<document.all.selFont.children.length;i++) if(document.all.selFont.children[i].value==photoObj.style.fontFamily) {document.all.selFont.selectedIndex=i;break;}
			document.all.selSize.disabled=false;
			for(var i=0;i<document.all.selSize.children.length;i++) if(document.all.selSize.children[i].value==photoObj.style.fontSize) {document.all.selSize.selectedIndex=i;break;}
			document.all.selColor.disabled=false;
			if(photoObj.style.display=='none') {
				document.all.selColor.selectedIndex=0;
			} else {
				for(var i=1;i<document.all.selColor.children.length;i++) if(document.all.selColor.children[i].value.toLowerCase()==photoObj.style.color.toLowerCase()) {document.all.selColor.selectedIndex=i;break;}
			}
			document.all.rot_deg[0].disabled=false;
			document.all.rot_deg[1].disabled=false;
			document.all.rot_deg[2].disabled=false;
			document.all.rot_deg[3].disabled=false;
			document.all.rot_deg[photoObj.filters.item(0).Rotation].checked=true;
			document.all.selBold.disabled=false;
			if(photoObj.style.fontWeight=="bolder") {
				document.all.selBold.checked=true;
			} else {
				document.all.selBold.checked=false;
			}
			document.all.selItalic.disabled=false;
			if(photoObj.style.fontStyle=="italic") {
				document.all.selItalic.checked=true;
			} else {
				document.all.selItalic.checked=false;
			}
		} else {
			document.onkeydown=new Function("");
			document.all.photo_mode[0].disabled=true;
			document.all.photo_mode[1].disabled=true;
			document.all.photo_mode[2].disabled=true;
			document.all.photo_mode[3].disabled=true;
			document.all.selFont.disabled=true;
			document.all.selFont.selectedIndex=0;
			document.all.selSize.disabled=true;
			document.all.selSize.selectedIndex=3;
			document.all.selColor.disabled=true;
			document.all.selColor.selectedIndex=137;
			document.all.rot_deg[0].disabled=true;
			document.all.rot_deg[1].disabled=true;
			document.all.rot_deg[2].disabled=true;
			document.all.rot_deg[3].disabled=true;
			document.all.selBold.disabled=true;
			document.all.selBold.checked=false;
			document.all.selItalic.disabled=true;
			document.all.selItalic.checked=false;
		}
	}

	function selectAccouterment(selValue) {
		document.all.SelName.selectedIndex=0;
		selName("");
		document.all.SelBody.selectedIndex=0;
		selBody("");
		selAccouterment(selValue);
	}
	
	function selAccouterment(selValue) {
		photoObj=document.all.ThePhoto;
		for(var i=0;i<photoObj.children.length;i++) 
			photoObj.children[i].style.borderStyle="";
		if(selValue!="") {
			document.onkeydown=process;
			photoObj=eval("document.all.AccoutermentBorder_"+selValue);
			photoObj.style.borderWidth="1px";
			photoObj.style.borderStyle="dotted";
			photoObj.style.borderColor="#ff0000";
			document.all.accouterment.disabled=false;
			for(var i=0;i<document.all.accouterment.children.length;i++) {
				if(eval("photo_design.AccoutermentUserVisual_"+document.all.SelAccouterment.options[document.all.SelAccouterment.selectedIndex].value).value==document.all.accouterment.children[i].value) {
					document.all.accouterment.selectedIndex=i;
					break;
				}
			}
			document.all.mirror.disabled=false;
			if(photoObj.filters.item(0).Mirror!=0) {
				document.all.mirror.checked=true;
			} else {
				document.all.mirror.checked=false;
			}
			document.all.photo_mode[0].disabled=false;
			document.all.photo_mode[1].disabled=false;
			document.all.photo_mode[2].disabled=false;
			document.all.photo_mode[3].disabled=false;
		} else {
			document.onkeydown=new Function("");
			document.all.accouterment.disabled=true;
			document.all.accouterment.selectedIndex=0;
			document.all.photo_mode[0].disabled=true;
			document.all.photo_mode[0].checked=true;
			document.all.photo_mode[1].disabled=true;
			document.all.photo_mode[2].disabled=true;
			document.all.photo_mode[3].disabled=true;
			document.all.mirror.disabled=true;
			document.all.mirror.checked=false;
		}
	}
	
	function selectFont(selValue) {
		photoObj.style.fontFamily=selValue;
	}

	function selectSize(selValue) {
		photoObj.style.fontSize=selValue;
		photoObj.style.lineHeight=selValue;
	}

	function selectColor(selValue) {
		if(selValue=="") {
			photoObj.style.display='none';
			photoObj.style.color='#ffffff';
		} else {
			photoObj.style.display='inline';
			photoObj.style.color=selValue;
		}
	}

	function selectBold() {
		if(photoObj.style.fontWeight=='bolder') {
			photoObj.style.fontWeight='normal';
		} else {
			photoObj.style.fontWeight='bolder';
		}
	}

	function selectItalic() {
		if(photoObj.style.fontStyle=='italic') {
			photoObj.style.fontStyle='normal';
		} else {
			photoObj.style.fontStyle='italic';
		}
	}

	function selectRotation(selValue) {
		photoObj.filters.item(0).rotation=selValue;
	}

	function selectMirror() {
		photoObj.filters.item(0).mirror=1-photoObj.filters.item(0).mirror;
	}

	function DoTitle(addTitle) {  
		var revisedTitle;  
		var currentTitle = document.photo_request.photouser.value; 
		if(currentTitle=="") revisedTitle = addTitle; 
		else { 
			var arr = currentTitle.split(","); 
			for (var i=0; i < arr.length; i++) { 
				if(addTitle.indexOf(arr[i]) >=0 && arr[i].length==addTitle.length ) return; 
			}
			revisedTitle = currentTitle+","+addTitle; 
		} 
		document.photo_request.photouser.value=revisedTitle;  
		document.photo_request.photouser.focus(); 
		return; 
	} 

	function selBg(selValue) {
		if(selValue!='') {
			if(selValue!='0') {
				document.all.ThePhoto.style.pixelWidth=280;
				document.photo_bg.style.pixelWidth=280;
				photo_request.bgwidth.value="280";
				photo_request.photobodyback.disabled=false;
				photo_request.photobodyfore.disabled=false;
				photo_request.photofg.disabled=false;
				document.all.ThePhoto.children[5].style.pixelLeft=70;
				document.all.ThePhoto.children[6].style.pixelLeft=105;
				if(selValue=="1" && photo_request.bgfname.value=="") {
					document.photo_bg.src='';
					document.photo_bg.style.display='none';
				} else {
					if(selValue!="1") {
						document.photo_bg.src="<%=PicPath%>"+selValue;
					} else {
						document.photo_bg.src="<%=PhotoPath%>/"+photo_request.bgfname.value;
					}
					document.photo_bg.style.display='inline';
				}
				document.photo_mask.src='images/img_visual/unfinished.gif';
				document.photo_mask.style.pixelWidth=280;
			} else {
				document.all.ThePhoto.style.pixelWidth=140;
				document.photo_bg.style.pixelWidth=140;
				document.photo_bodyback.src='';
				document.photo_bodyback.style.display='none';
				photo_request.photobodyback.disabled=true;
				document.photo_bodyfore.src='';
				document.photo_bodyfore.style.display='none';
				photo_request.photobodyfore.disabled=true;
				document.photo_fg.src='';
				document.photo_fg.style.display='none';
				photo_request.photofg.disabled=true;
				photo_request.bgwidth.value="140";
				document.all.ThePhoto.children[5].style.pixelLeft=0;
				document.all.ThePhoto.children[6].style.pixelLeft=35;
				document.photo_bg.src="<%=PicPath%>"+MyBg;
				if(MyBg!="") document.photo_bg.style.display='inline';
				document.photo_mask.src='images/img_visual/unfinished_s.gif';
				document.photo_mask.style.pixelWidth=140;
			}
		} else {
			document.all.ThePhoto.style.pixelWidth=280;
			document.photo_bg.style.pixelWidth=280;
			photo_request.bgwidth.value="280";
			photo_request.photobodyback.disabled=false;
			photo_request.photobodyfore.disabled=false;
			photo_request.photofg.disabled=false;
			document.all.ThePhoto.children[5].style.pixelLeft=70;
			document.all.ThePhoto.children[6].style.pixelLeft=105;
			document.photo_bg.src='';
			document.photo_bg.style.display='none';
		}
	}

	function selBodyBack(selValue) {
		if(selValue!='') {
			document.photo_bodyback.src="<%=PicPath%>"+selValue;
			document.photo_bodyback.style.display='inline';
		} else {
			document.photo_bodyback.src='';
			document.photo_bodyback.style.display='none';
		}
	}

	function selBodyFore(selValue) {
		if(selValue!='') {
			document.photo_bodyfore.src="<%=PicPath%>"+selValue;
			document.photo_bodyfore.style.display='inline';
		} else {
			document.photo_bodyfore.src='';
			document.photo_bodyfore.style.display='none';
		}
	}

	function selFg(selValue) {
		if(selValue!='') {
			document.photo_fg.src="<%=PicPath%>"+selValue;
			document.photo_fg.style.display='inline';
		} else {
			document.photo_fg.src='';
			document.photo_fg.style.display='none';
		}
	}

	function setAccouterment(selValue) {
		if(selValue!='') {
			photoObj.children[0].src="<%=PicPath%>"+selValue;
			photoObj.children[0].style.display='inline';
			eval("photo_design.AccoutermentUserVisual_"+document.all.SelAccouterment.options[document.all.SelAccouterment.selectedIndex].value).value=selValue;
		} else {
			photoObj.children[0].src='';
			photoObj.children[0].style.display='none';
			eval("photo_design.AccoutermentUserVisual_"+document.all.SelAccouterment.options[document.all.SelAccouterment.selectedIndex].value).value="";
		}
	}

	function SubmitDesign() {
		photo_design.PhotoNameLeft.value=document.all.PhotoNameBorder.style.pixelLeft;
		photo_design.PhotoNameTop.value=document.all.PhotoNameBorder.style.pixelTop;
		photo_design.PhotoNameFont.value=document.all.PhotoNameBorder.style.fontFamily;
		photo_design.PhotoNameSize.value=document.all.PhotoNameBorder.style.fontSize;
		if(eval("document.all.PhotoNameBorder").style.fontWeight=="bolder") {
			eval("photo_design.PhotoNameBold").value=1;
		} else {
			eval("photo_design.PhotoNameBold").value=0;
		}
		if(eval("document.all.PhotoNameBorder").style.fontStyle=="italic") {
			eval("photo_design.PhotoNameItalic").value=1;
		} else {
			eval("photo_design.PhotoNameItalic").value=0;
		}
		if(document.all.PhotoNameBorder.style.display=="none") {
			photo_design.PhotoNameColor.value="";
		} else {
			photo_design.PhotoNameColor.value=document.all.PhotoNameBorder.style.color;
		}
		eval("photo_design.PhotoNameDirection").value=eval("document.all.PhotoNameBorder").filters.item(0).rotation;
		photo_design.DateLeft.value=document.all.DateBorder.style.pixelLeft;
		photo_design.DateTop.value=document.all.DateBorder.style.pixelTop;
		photo_design.DateFont.value=document.all.DateBorder.style.fontFamily;
		photo_design.DateSize.value=document.all.DateBorder.style.fontSize;
		if(eval("document.all.DateBorder").style.fontWeight=="bolder") {
			eval("photo_design.DateBold").value=1;
		} else {
			eval("photo_design.DateBold").value=0;
		}
		if(eval("document.all.DateBorder").style.fontStyle=="italic") {
			eval("photo_design.DateItalic").value=1;
		} else {
			eval("photo_design.DateItalic").value=0;
		}
		if(document.all.DateBorder.style.display=="none") {
			photo_design.DateColor.value="";
		} else {
			photo_design.DateColor.value=document.all.DateBorder.style.color;
		}
		eval("photo_design.DateDirection").value=eval("document.all.DateBorder").filters.item(0).rotation;
		for(var i=0;i<userCount;i++) {
			eval("photo_design.LayerNo_"+i).value=eval("document.all.BodyBorder_"+i).style.zIndex;
			eval("photo_design.OuterLeft_"+i).value=eval("document.all.BodyBorder_"+i).style.pixelLeft;
			eval("photo_design.OuterTop_"+i).value=eval("document.all.BodyBorder_"+i).style.pixelTop;
			eval("photo_design.OuterWidth_"+i).value=eval("document.all.BodyBorder_"+i).style.pixelWidth;
			eval("photo_design.OuterHeight_"+i).value=eval("document.all.BodyBorder_"+i).style.pixelHeight;
			eval("photo_design.InnerLeft_"+i).value=eval("document.all.BodyCont_"+i).style.pixelLeft;
			eval("photo_design.InnerTop_"+i).value=eval("document.all.BodyCont_"+i).style.pixelTop;
			eval("photo_design.InnerWidth_"+i).value=eval("document.all.BodyCont_"+i).style.pixelWidth;
			eval("photo_design.InnerHeight_"+i).value=eval("document.all.BodyCont_"+i).style.pixelHeight;
			eval("photo_design.Direction_"+i).value=eval("document.all.BodyBorder_"+i).filters.item(0).mirror*4;
			eval("photo_design.NameLeft_"+i).value=eval("document.all.NameBorder_"+i).style.pixelLeft;
			eval("photo_design.NameTop_"+i).value=eval("document.all.NameBorder_"+i).style.pixelTop;
			eval("photo_design.NameFont_"+i).value=eval("document.all.NameBorder_"+i).style.fontFamily;
			eval("photo_design.NameSize_"+i).value=eval("document.all.NameBorder_"+i).style.fontSize;
			if(eval("document.all.NameBorder_"+i).style.fontWeight=="bolder") {
				eval("photo_design.NameBold_"+i).value=1;
			} else {
				eval("photo_design.NameBold_"+i).value=0;
			}
			if(eval("document.all.NameBorder_"+i).style.fontStyle=="italic") {
				eval("photo_design.NameItalic_"+i).value=1;
			} else {
				eval("photo_design.NameItalic_"+i).value=0;
			}
			if(eval("document.all.NameBorder_"+i).style.display=="none") {
				eval("photo_design.NameColor_"+i).value="";
			} else {
				eval("photo_design.NameColor_"+i).value=eval("document.all.NameBorder_"+i).style.color;
			}
			eval("photo_design.NameDirection_"+i).value=eval("document.all.NameBorder_"+i).filters.item(0).rotation;
		}
		for(var i=0;i<6;i++) {
			photoObj=eval("document.all.AccoutermentBorder_"+i);
			eval("photo_design.AccoutermentOuterLeft_"+i).value=photoObj.style.pixelLeft;
			eval("photo_design.AccoutermentOuterTop_"+i).value=photoObj.style.pixelTop;
			eval("photo_design.AccoutermentOuterWidth_"+i).value=photoObj.style.pixelWidth;
			eval("photo_design.AccoutermentOuterHeight_"+i).value=photoObj.style.pixelHeight;
			eval("photo_design.AccoutermentInnerLeft_"+i).value=photoObj.children[0].style.pixelLeft;
			eval("photo_design.AccoutermentInnerTop_"+i).value=photoObj.children[0].style.pixelTop;
			eval("photo_design.AccoutermentInnerWidth_"+i).value=photoObj.children[0].style.pixelWidth;
			eval("photo_design.AccoutermentInnerHeight_"+i).value=photoObj.children[0].style.pixelHeight;
			eval("photo_design.AccoutermentDirection_"+i).value=photoObj.filters.item(0).mirror*4;
		}
	}
</script>