var i_messages = 0
var timer

function dotransition() {
	
    if (document.all) {
        content.filters[0].apply()
		
        content.innerHTML = "<font color="+mescolor[i_messages]+">"+messages[i_messages]+"</font>"
        content.filters[0].play()
        if (i_messages >= messages.length-1) {
            i_messages = 0
        }
        else {
            i_messages++
        }
    } 
    timer = setTimeout("dotransition()",6000)  
}