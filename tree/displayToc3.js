/*------------------------------------------------------------------- 
Author's 
Statement: This script is based on ideas of the author. You may copy, modify and 
use it for any purpose. The only condition is that if you publish web pages that 
use this script you point to its author at a suitable place and don't remove 
this Statement from it. It's your responsibility to handle possible bugs even if 
you didn't modify anything. I cannot promise any support. Dieter Bungers GMD 
(www.gmd.de) and infovation (www.infovation.de) 
--------------------------------------------------------------------*/



if (navigator.appName.toLowerCase().indexOf("explorer") > -1) {
	var mdi=textSizes[1], sml=textSizes[2], soff=21;
}
else {
	var mdi=textSizes[3], sml=textSizes[4], soff=18;
}

function reDisplay(currentNumber,currentIsExpanded) {
	toc.document.open();
	toc.document.write("<html>\n<head>\n<title>ToC</title>\n<style TYPE=text/css>\n")
	toc.document.write (".ftoc {font-family:'MS Sans Serif',Helvetica,Arial;font-weight:normal;font-size:0.7em;color: #000000;text-decoration:none;}\n")
	toc.document.write (".btoc {font-family:'MS Sans Serif',Helvetica,Arial;font-weight:normal;font-size:0.7em;}\n")
	toc.document.write (".btoc A:hover{text-decoration:underline;}\n")
	toc.document.write("</style></head>\n<body bgcolor=\"#FFFFFF\" link=#0000CC vlink=#000077>\n<table border=0 cellspacing=1 cellpadding=0>\n<tr>");
	var currentNumArray = currentNumber.split(".");
	var currentLevel = currentNumArray.length-1;
	
	var scrollY=0, addScroll=true, theHref=""; 
	for (i=0; i<tocTab.length; i++) {
		thisNumber = tocTab[i][0];
		var isCurrentNumber = (thisNumber == currentNumber);
		//if (isCurrentNumber) theHref=tocTab[i][2];
		theHref = tocTab[i][2];
		var thisNumArray = thisNumber.split(".");
		var thisLevel = thisNumArray.length-1;
		var toDisplay = true;
		if (thisLevel > 0) {
			
			for (j=0; j<thisLevel; j++) {
				toDisplay = (j>currentLevel)?false:toDisplay && (thisNumArray[j] == currentNumArray[j]);
			}
		}
		thisIsExpanded = toDisplay && (thisNumArray[thisLevel] == currentNumArray[thisLevel])
		if (currentIsExpanded) {
			toDisplay = toDisplay && (thisLevel<=currentLevel);
			if (isCurrentNumber) thisIsExpanded = false;
		}
		
		if (toDisplay) {
			if (i==0) {
				
				toc.document.writeln("\n<td colspan=" + (nCols+1) + "><small> </td></tr>");
				for (k=0; k<nCols; k++) {
					toc.document.write("<td><small> </td>");
				}
				toc.document.write("<td width=1234><small> </td></tr>");
			}
			
			else {
				if (addScroll) scrollY+=((thisLevel<2)?mdi:sml)*soff
				if (isCurrentNumber) addScroll=false;
				var isLeaf = (i==tocTab.length-1) || (thisLevel >= tocTab[i+1][0].split(".").length-1);
				img = (isLeaf)?"a":(thisIsExpanded)?"o":"f";
				
				toc.document.writeln("<tr>");
				for (k=1; k<=thisLevel; k++) {
					toc.document.writeln("<td><small> </td>");
				}
				// must add target='_content' to the anchor, change the javascript 
				// function to an onclick, and change the href to theHref in the 
				// following line for this to work in a netscape "my sidebar" tab
				//toc.document.writeln("<td valign=top><a href=\"javaScript:parent.reDisplay('" + thisNumber + "'," + thisIsExpanded + ");\"><img src=\"" + img + ".gif\" border=0></a></td> <td colspan=" + (nCols-thisLevel) + "><a href=\"javaScript:\parent.reDisplay('" + thisNumber + "'," + thisIsExpanded + ")\" style=\"font-family: " + fontLines + ";" + ((thisLevel<=mLevel)?"font-weight:normal":"") +  "; font-size:" + ((thisLevel<=mLevel)?mdi:sml) + "em; color: " + ((isCurrentNumber)?currentColor:normalColor) + "; text-decoration:none\">" + ((showNumbers)?(thisNumber+" "):"") + "<nobr>" + ((bold)?"<b>":" ") + tocTab[i][1] + ((bold)?"</b>":" ") + "</a></td></tr>");
				if (theHref != "")  // something about this part messes up the href:  onClick=\"parent.reDisplay('" + thisNumber + "'," + thisIsExpanded + ");\"
					toc.document.writeln("<td valign=top><a href=\"" + theHref +  "\" target=\"content\"><img src=\"" + img + ".gif\" border=0></a></td> <td colspan=" + (nCols-thisLevel) + "><a href=\"" + theHref +  "\" target=\"content\" class=\"btoc\"><nobr>" + tocTab[i][1] + "</a></td></tr>");
				else 
					toc.document.writeln("<td valign=top><a href=\"javaScript:parent.reDisplay('" + thisNumber + "'," + thisIsExpanded + ");\"><img src=\"" + img + ".gif\" border=0></a></td> <td colspan=" + (nCols-thisLevel) + "><a href=\"javaScript:\parent.reDisplay('" + thisNumber + "'," + thisIsExpanded + ");\" class=\"ftoc\"><nobr> " + tocTab[i][1] + "</a></td></tr>");
			}
		}
	}
	toc.document.writeln("</table>\n");
	//toc.document.writeln("<a href=\"\" onClick=\"javascript:parent.tocTab[0][1] = 'changed it back'; return false;\"><img src= \"a.gif\"></a>");
	toc.document.writeln("\n</body>");
	toc.document.close();
	toc.scroll(0,scrollY);
	
	// here is where it redirects the page
	//if (theHref != "") _content.location.href = theHref;
	//return true;
}