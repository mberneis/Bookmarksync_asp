<?php

header("Content-type: text/plain");

//*****************************************************************
//*
//* SYNCIT.php
//*
//* Last Modified: 10/2/99
//*
//* Revision History
//* 17/11/98	mb	Added Comments
//* 18/11/98	tw	Changed subscription loop (remove unnecessary
//*         		test) and put email address on subscription name
//* 29/12/98     tw      Removed *M, changed *V for 0.15
//* 10/2/99	mb	Changes for Publisher/Subscriptions Model
//*			Locking changed
//*			Links get expiration
//* 10/3/99	tw	create/remove directories, allow same path,
//*			different URLs for expired bookmarks
//*****************************************************************

//*** Will be set to true if at least one delete or add statement
$changed = false;


$debug = true;		// set this to true to log all sql statements in parseline()



//*********************************************** MAIN PROGRAM *********************************************

//*** Get Parameter
$email = trim($_POST["Email"]);
$pass = trim($_POST["md5"]);
$version = trim($_POST["Version"]);
$content = $_POST["Contents"];
$ctoken = intval($_POST["Token"]);
$charset = trim($_POST['CharSet']);

////////////////////////////////////////////////////////////////////////
// debug code to see what the client sends
////////////////////////////////////////////////////////////////////////
if ($debug == true){
	ob_start();
	print_r($_POST);
	$str = ob_get_contents();
	ob_end_clean();
	$fp = fopen ("client_log.txt","a");
	fputs ($fp, $str . "\r\n");
	fclose($fp);
}
////////////////////////////////////////////////////////////////////////


//* Open Database
//mysql_connect("mysql.syncitlinux.com","synci001","glaister");
//mysql_select_db("synci001");
mysql_connect("localhost","fluffy","armani");
mysql_select_db("syncit");

// *** check email address ***
$res = my_mysql_query("select personid,token,pass from person where email='" . $email . "'");
$data = mysql_fetch_assoc($res);

if (!$data){
	echo "*N\r\n*Z\r\n";
	exit();
}
else {
	$ID = $data["personid"];
	$stoken = intval($data["token"]);
}

// *** check password ***
// generate md5 hash first
$md5pw = base64_encode(pack("H*",md5($email . $data['pass'])));
if ($md5pw != $pass) {
	echo "*P\r\n*Z\r\n";
	exit();
}

echo "*S,Root,\"http://" . $_SERVER['SERVER_NAME'] . "/\"\r\n";

// *** compare tokens ***
if ($stoken > $ctoken){
	// return bookmark list
	$res = my_mysql_query("select path,url,o.src as io,c.src as ic from link left join images as o on o.imgid = openimg_id 
						left join images as c on c.imgid = closeimg_id, bookmarks where bookid=book_id and person_id=" . $ID . " and expiration is NULL order by path");

	echo "*T," . $stoken . "\r\n*B\r\n";
	while ($data = mysql_fetch_assoc($res))
		echo "\"" . $data["path"] . "\",\"" . $data["url"] . "\"\r\n";	//," . $data["io"] . "," . $data["ic"] . "\r\n";
}
else {
	//*** Bookmarklist seems uptodate - lets process A/D directives now
	if (strlen($content) > 0){
		$changed = false;
		do {
			//*** Go through the A/D commandlist extract line by line and call parseline
			$i = strpos($content,"\r\n");
			if ($i == 0)
				break;
			
			$bm_row = substr($content,0,$i);
			if (parseline($bm_row,$ID)){
				echo "*Z\r\n";
				exit();
			} 
			
			$content = substr($content,strlen($content)-(strlen($content)-$i-1));
		} while (true);

		//conn.BeginTrans);
		my_mysql_query("update person set token = token+1,lastchanged=now() where personid=" . $ID);
		$res = my_mysql_query("select token from person where personid=" . $ID);
		$data = mysql_fetch_assoc($res);
		$stoken = $data["token"];
		//conn.CommitTrans

		if ($changed)
			//*** Mark all publications with current token
			my_mysql_query("update publish set token = " . $stoken . " where user_id=" . $ID);
	} 

	//*** Return new token
	echo "*T," . $stoken . "\r\n";
}

/*
// Return Subscriptions
$count = 0;
$res = my_mysql_query("select publishid,title,user_id,path,token from publish,subscriptions where subscriptions.person_id=" . $ID . " and subscriptions.publish_id=publish.publishid order by title");
while ($data = mysql_fetch_assoc($res)){

	$rtoken = $data["token"];
	$rid = $data["publishid"];
	$plen = strlen($data["path"])+1;					// will add leading backslash soon
	if ($_POST["token".$rid][$count] ==1 && intval($_POST["token".$rid]) == $rtoken)
		echo "*Q," . $rid . ",\"" . $data["title"] . "\"," . $rtoken . "\r\n";
	else {
		if (substr($data["title"],0,7) == "Gartner")
			echo "*R," . $rid . ",\"" . $data["title"] . "\"," . $rtoken . ",icons/gartner.bmp\r\n";
		else
			echo "*R," . $rid . ",\"" . $data["title"] . "\"," . $rtoken . "\r\n";
	
		$res1 = my_mysql_query("select path,url,o.src AS io,c.src AS ic FROM link LEFT JOIN images AS o ON o.imgid = openimg_id 
								LEFT JOIN images AS c ON c.imgid = closeimg_id,bookmarks WHERE bookid=book_id and person_id= "
								. $data["user_id"] . "AND path LIKE '\\" . $data["path"] . "\%') AND expiration IS NULL ORDER BY path");
		while($data1 = mysql_fetch_assoc($res1)){
			$path = $data1["path"];
			echo "\"" . substr($path,strlen($path)-(strlen($path)-$plen)) . "\",\"" . $data1["url"] . "\"," . $data1["io"] . "," . $data1["ic"] . "\r\n";
		} 
	}
} 
*/

//*** Close up
echo "*Z\r\n";
exit();




//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//* parses bookmarklist line, processes Adds and Deletes
function parseline($bm_row,$ID)
{
	// sheesh!
//	$bm_row = stripslashes($bm_row);		// this was the culprit!!!!!!! duh!!
	$bm_row = trim($bm_row);

	$cmd = substr($bm_row,0,1);												//*** first char is command
	$bm_row = substr($bm_row,3);			// *** strip 1st token
	$l = strpos($bm_row,"\"");												//*** find "
	$tpath = substr($bm_row,0,$l);										// *** get path
	$path = str_replace("'","''",$tpath);									// *** SQL escape single quotes

	// add bookmark
	if ($cmd == "A"){
		$url = substr($bm_row,strlen($tpath)+3);
		$url = trim($url,"\"\r\n");

		$res = my_mysql_query("select bookid from bookmarks where url='" . $url . "'");
		if (mysql_num_rows($res) == 0){
			my_mysql_query("insert into bookmarks (url) Values ('" . $url . "')");
			$bid = mysql_insert_id();
		}
		else {
			$data = mysql_fetch_assoc($res);
			$bid = $data["bookid"];
		}
		$res = my_mysql_query("insert into link (expiration,person_id,access,path,book_id) values (NULL," . $ID . ",now(),'" . addslashes($path) . "'," .$bid . ")");
		// Bookmark exists already - but expired?: Unexpire!
		if (!$res)
			my_mysql_query("update link set expiration = NULL,book_id = " . $bid . " where person_id = " . $ID . " and path = '" . addslashes($path) . "'");

	}

	// delete it = set expiration
	else if ($cmd == "D")
		my_mysql_query("update link set expiration = now() where path = '\\" . $path . "' and person_id = " . $ID);

	// make directory
	else if ($cmd == "M"){
		$res = my_mysql_query("insert into link (expiration,person_id,access,path) values (NULL," . $ID . ",now(),'" . addslashes($path) . "')");
		if (!$res)
			my_mysql_query("update link set expiration = NULL, book_id = NULL where person_id = " . $ID . " and path = '" . addslashes($path) . "'");
	}

	// remove directory
	else if ($cmd == "R")
		my_mysql_query("update link set expiration = now() where path = '" . addslashes($path) . "' and person_id = " . $ID);

	// invalid command
	else {
		echo "*E\r\nInvalid Bookmark Command: " .$cmd . "\r\n";
		return true;
	}

	$changed = true;
	return false;
}


function my_mysql_query($str)
{
	global $debug;

	if ($debug == true){
		$fp = fopen("my_mysql.log","a");
		fputs($fp,$str . "\r\n");
		fclose($fp);
	}

	$res = mysql_query($str);
	return $res;
}

?>