<?php
// (c) 2003-2004 Stealth Networks all rights reserved
// Directs the user to a randomly-chosen StealthBot mirror :)
// Updated February 2006 to match the new website skin Minimalist
function checkOutURL()
{
    $outsite = getData('out', 'G', 'INT');

    if( $outsite > 0 )
    {
        //logAccess();

        getMirror($outsite, $full, $chosen, $outurl, $imgurl);
                                    
        header("Location: " . $outurl);
        die();
    }
}

function logAccess()
{
    define('_VALID_CODE', true);
    include_once('database.php');
    $dbh = new Database();
    $dbh->init();

    $ipaddr_base = $_SERVER['REMOTE_ADDR'];
    if( is_array($ipaddr_base) )
        $ipaddr = $ipaddr_base[0];
    else
        $ipaddr = $ipaddr_base;


    //echo("--> " . $ipaddr);

    $query = "SELECT * FROM downloads26 WHERE ipaddr = '$ipaddr';";
    $result = $dbh->doQuery($query);

    if(mysql_num_rows(&$result) > 0)
    {
        /*$row = mysql_fetch_array(&$result);

        if (intval( $row['attempts'] ) > 50)
        {
            die("Sorry, you've attempted to download StealthBot too many times. Please contact Stealth by <a href='mailto:stealth@stealthbot.net'>email</a> if you need another copy, or if you received this message in error.");
        } */

        $query = "UPDATE downloads26 SET attempts=attempts+1 WHERE ipaddr='$ipaddr';";
    }
    else
    {
        $query = "INSERT INTO downloads26(ipaddr, attempts) VALUES('$ipaddr', 1);";
    }

    mysql_free_result($result);

    $result = $dbh->doQuery($query);
}

function randomMirror ()
{
	//echo("Currently uploading a new version. Check back in 10-45 minutes.");
	//return;

	$MIRROR_COUNT = 5;

	$rn = rand(1, $MIRROR_COUNT);

	$att = getData('select', 'G', 'INT');

	if( is_numeric($att) )
	{
		if( ! (($att > $MIRROR_COUNT) || ($att < 1)) )
		{
			$rn = $att;
		}
	}

	$version = "2.6 Revision 3";

    getMirror($rn, $full, $chosen, $outurl, $imgurl);

	print ("<center><br><br>");
	print ("Thank you for choosing StealthBot. "); // You have been redirected to mirror $rn.<br>");
	//print ("This mirror is hosted by <a href=\"http://$chosen\" target=\"_blank\">$full</a>.<br><br>");
	print ("You are downloading version <b>$version</b><br><br>");
	print ("<a href=\"http://$chosen\" target=\"_blank\"><img src=\"$imgurl\" alt=\"$full\"></a><br><br>");
	print ("<strong><a href=\"getsb.php?out=$rn\">Click here to begin your download!</a></strong><br><br>Note that downloading, installing and operating StealthBot is subject to the terms of the <a href=\"http://eula.stealthbot.net\">End-User License Agreement</a>.<br>");
	//print ("<br /><span style='font weight: bold'>Please check out our mirrors if you get a chance, they've kindly donated their bandwidth so you can have the latest StealthBot!</span></center><br /><br />");

}

function getMirror($num, &$full, &$chosen, &$outurl, &$imgurl)
{
    $fn = "InstallSB.exe";

    /*
    switch ($num)
    {
        // double coverage for me
        case 1:
        case 4:
        case 5:*/
            $full = "StealthBot.net";
            $chosen = "stealthbot.net";
            $outurl = "http://stealthbot.net/dist/" . $fn;
            $imgurl = "http://www.stealthbot.net/forum/style_images/2/logo4.gif";
            //break;

     /*
        default:
            die("Invalid mirror!");
    }  */
    
}


// getData() retrieves GET or POST variables
// origins in LDU 604 core
function getData($varname, $source, $out_type)
{
	switch($source)
	{
		case 'G':
    		$cv_o = $_GET[$varname];
    		break;

		case 'P':
    		$cv_o = $_POST[$varname];
    		break;

		case 'C':
    		$cv_o = $_COOKIE[$varname];
    		break;

		default:
		    die ("Unknown source for a variable.");
    		break;
	}

	if ($cv_o=='')
       	{ return(''); }

	switch($out_type)
	{
		case 'INT':
    		if (is_numeric($cv_o)==TRUE && floor($cv_o)==$cv_o)
    	    {
                return($cv_o);
            }
    	    else
    		{
    		    return(NULL);
    		}
    		break;

		case 'NUM':
		if (is_numeric($cv_o)==TRUE)
	       	{
                return($cv_o);
            }
	       else
		   {
			    return(NULL);
		   }
		   break;

		case 'TXT':
    		if (strpos($cv_o,"<")===FALSE)
    		{
                return(strip_tags($cv_o));
            }
    	    else
    		{
    		    return(NULL);
    		}
    		break;

		case 'BLN':
    		if (is_bool($cv_o)==TRUE)
    	    {
                return($cv_o);
            }
    	    elseif ($cv_o=="0")
    	    {
                return(FALSE);
            }
    	    elseif ($cv_o=="1" || $cv_o=="on")
    	    {
                return(TRUE);
            }
    	    else
    		{
    			return(NULL);
    		}
    		break;

		case 'LVL':
    		if (is_numeric($cv_o)==TRUE && $cv_o>-1 && $cv_o<101 && floor($cv_o)==$cv_o)
    	    {
                return($cv_o);
            }
    	    else
    		{
    			return(NULL);
    		}
    		break;

		case 'NOC':
    		return($cv_o);
    		break;

		default:
    		die("Unknown type for a variable : <br />Var = ".$varname."<br />Type = ".$out_type." ?");
    		break;
	}
}


// Begin document output

checkOutURL();
?>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xml:lang="en" xmlns="http://www.w3.org/1999/xhtml" lang="en">

<head>
<title>StealthBot.net -> Download StealthBot</title> 

<meta http-equiv="content-type" content="text/html; charset=iso-8859-1"> 
<style type="text/css">
html { overflow-x: auto; }
BODY { font-family: Tahoma, Verdana, Arial, sans-serif;font-size: 11px;margin: 0px;padding: 0px;text-align: center;color: #000;background-color: #DDDDDD; }
TABLE, TR, TD { font-family: Tahoma, Verdana, Arial, sans-serif;font-size: 11px;color: #000;background-color: #FFFFFF; }
.small { font-size: 10px; }
.toplinks { font-size: 10px;font-weight: bold;color: #FFFFFF; }
.toplinks a:link, .toplinks a:visited, .toplinks a:hover, .toplinks a:hover { text-decoration: none;color: #FFFFFF; }
.brackets { color: #000000; }
.brackets a:link, .brackets a:hover, .brackets a:active, .brackets a:visited { text-decoration: none;color: #000000; }
#ipbwrapper { text-align: left;width: 95%;margin-left: auto;margin-right: auto; }
a:link, a:visited, a:active { text-decoration: none;color: #1E4166; }
a:hover { text-decoration: none;color: #000000; }
fieldset.search { padding: 6px;line-height: 150%; }
label { cursor: pointer; }
form { display: inline; }
img { vertical-align: middle;border: 0px; }
img.attach { padding: 2px;border: 2px outset #F7F7F7; }
.googleroot { padding: 6px;line-height: 130%; }
.googlechild { padding: 6px;margin-left: 30px;line-height: 130%; }
.googlebottom, .googlebottom a:link, .googlebottom a:visited, .googlebottom a:active { font-size: 11px;color: #3A4F6C; }
.googlish, .googlish a:link, .googlish a:visited, .googlish a:active { font-size: 14px;font-weight: bold;color: #00D; }
.googlepagelinks { font-size: 1.1em;letter-spacing: 1px; }
.googlesmall, .googlesmall a:link, .googlesmall a:active, .googlesmall a:visited { font-size: 10px;color: #434951; }
li.helprow { padding: 0px;margin: 0px 0px 10px 0px; }
ul#help { padding: 0px 0px 0px 15px; }
option.cat { font-weight: bold; }
option.sub { font-weight: bold;color: #555; }
.caldate { text-align: right;font-weight: bold;font-size: 11px;padding: 4px;margin: 0px;color: #777;background-color: #E7E7E7; }
.warngood { color: green; }
.warnbad { color: red; }
#padandcenter { margin-left: auto;margin-right: auto;text-align: center;padding: 14px 0px 14px 0px; }
#profilename { font-size: 28px;font-weight: bold; }
#calendarname { font-size: 22px;font-weight: bold; }
#photowrap { padding: 6px; }
#phototitle { font-size: 24px;border-bottom: 1px solid black; }
#photoimg { text-align: center;margin-top: 15px; }
#ucpmenu { line-height: 150%;width: 22%;background-color: #FFFFFF;border: 1px solid #000000; }
#ucpmenu p { padding: 2px 5px 6px 9px;margin: 0px; }
#ucpcontent { line-height: 150%;width: auto;background-color: #FFFFFF;border: 1px solid #000000; }
#ucpcontent p { padding: 10px;margin: 0px; }
#ipsbanner { position: absolute;top: 1px;right: 5%; }
#logostrip { padding: 0px;margin: 0px;background-color: #3860BB;border: 1px solid #000000;background-image: url(style_images/Minimali-547/tile_back.gif); }
#submenu { font-size: 10px;margin: 3px 0px 3px 0px;font-weight: bold;color: #3A4F6C;background-color: #FFFFFF;border: 1px solid #000000; }
#submenu a:link, #submenu  a:visited, #submenu a:active { font-weight: bold;font-size: 10px;text-decoration: none;color: #3A4F6C; }
#userlinks { background-color: #FFFFFF;border: 1px solid 000000; }
#navstrip { font-weight: bold;padding: 6px 0px 6px 0px; }
.activeuserstrip { font-weight: bold;padding: 6px;margin: 0px;color: #141414;background-color: #EDEDED; }
.pformstrip { font-weight: bold;padding: 7px;color: #141414;background-color: #F0F0F0; }
.pformleft { padding: 6px;width: 25%;background-color: #FFFFFF; }
.pformleftw { padding: 6px;width: 40%;background-color: #FFFFFF; }
.pformright { padding: 6px;background-color: #FFFFFF; }
.signature { font-size: 10px;line-height: 150%;color: #141414; }
.postdetails { font-size: 10px; }
.postcolor { font-size: 12px;line-height: 160%; }
.normalname { font-size: 12px;font-weight: bold;color: #003; }
.normalname a:link, .normalname a:visited, .normalname a:active { font-size: 12px; }
.unreg { font-size: 11px;font-weight: bold;color: #900; }
.post1 { background-color: #FFFFFF; }
.post2 { background-color: #FFFFFF; }
.postlinksbar { padding: 7px;font-size: 10px;background-color: #E6E6E6; }
.row1 { background-color: #FFFFFF; }
.row2 { width: 10%;background-color: #FFFFFF; }
.row3 { background-color: #FFFFFF; }
.row4 { background-color: #FFFFFF; }
.darkrow1 { color: #141414;background-color: #F0F0F0; }
.darkrow2 { color: #141414;background-color: #EDEDED; }
.darkrow3 { color: #141414;background-color: #E9E9E9; }
.hlight { background-color: #FFFFFF; }
.dlight { background-color: #FFFFFF; }
.titlemedium { font-weight: bold;padding: 5px;margin: 0px;color: #000000;background-color: #EEEEEE; }
.titlemedium  a:link, .titlemedium  a:visited, .titlemedium  a:active { text-decoration: none;color: #141414; }
.titlemedium a:hover { text-decroration: none;color: #000000; }
.maintitle { vertical-align: middle;font-weight: bold;letter-spacing: 1px;padding: 6px 0px 6px 5px;color: #FFFFFF;background-color: #204873; }
.maintitle a:link, .maintitle  a:visited, .maintitle  a:active { text-decoration: none;color: #FFFFFF; }
.maintitle a:hover { text-decoration: none;color: #F0F0F0; }
.plainborder { background-color: #FFFFFF;border: 1px solid #000000; }
.tableborder { background-color: #FFF;border: 1px solid #000000; }
.tablefill { padding: 6px;background-color: #FFFFFF;border: 1px solid #000000; }
.tablepad { padding: 6px;background-color: #FFFFFF; }
.tablebasic { width: 100%;padding: 0px 0px 0px 0px;margin: 0px;background-color: #FFFFFF;border: 0px; }
.wrapmini { float: left;line-height: 1.5em;width: 25%; }
.pagelinks { float: left;line-height: 1.2em;width: 35%; }
.desc { font-size: 10px;color: #141414; }
.edit { font-size: 9px; }
.searchlite { font-weight: bold;color: #F00;background-color: #FF0; }
#QUOTE { font-family: Tahoma, Verdana, Arial;font-size: 11px;padding-top: 2px;padding-right: 2px;padding-bottom: 2px;padding-left: 2px;color: #141414;background-color: #FFFFFF;border: 1px solid #E6E6E6; }
#CODE { font-family: Courier, Courier New, Verdana, Arial;font-size: 11px;padding-top: 2px;padding-right: 2px;padding-bottom: 2px;padding-left: 2px;color: #141414;background-color: #FFFFFF;border: 1px solid #E6E6E6; }
.copyright { font-family: Tahoma, Verdana, Arial, Sans-Serif;font-size: 9px;line-height: 12px; }
.codebuttons { font-size: 10px;font-family: tahoma, verdana, helvetica, sans-serif;vertical-align: middle; }
.forminput, .textinput, .radiobutton, .checkbox { font-size: 11px;font-family: tahoma, verdana, helvetica, sans-serif;vertical-align: middle; }
.thin { padding: 2px 0px 2px 0px;margin: 2px 0px 2px 0px; }
.purple { font-weight: bold;color: purple; }
.red { font-weight: bold;color: red; }
.green { font-weight: bold;color: green; }
.blue { font-weight: bold;color: blue; }
.orange { font-weight: bold;color: #F90; }

</style></head><body>

<div id="ipbwrapper">

<!-- Minimalist skin (c) Ava for CalliopeDesigns.net 2004 -->
</div><table style="border: 1px solid rgb(0, 0, 0);" align="center" border="0" cellpadding="0" cellspacing="0" width="770">

<script language="JavaScript" type="text/javascript">
<!--
function buddy_pop() { window.open('index.php?act=buddy&s=','BrowserBuddy','width=250,height=500,resizable=yes,scrollbars=yes'); }
function chat_pop(cw,ch)  { window.open('index.php?s=&act=chat&pop=1','Chat','width='+cw+',height='+ch+',resizable=yes,scrollbars=yes'); }
function multi_page_jump( url_bit, total_posts, per_page )
{
pages = 1; cur_st = parseInt(""); cur_page  = 1;
if ( total_posts % per_page == 0 ) { pages = total_posts / per_page; }
 else { pages = Math.ceil( total_posts / per_page ); }
msg = "Please enter a page number to jump to between 1 and" + " " + pages;
if ( cur_st > 0 ) { cur_page = cur_st / per_page; cur_page = cur_page -1; }
show_page = 1;
if ( cur_page < pages )  { show_page = cur_page + 1; }
if ( cur_page >= pages ) { show_page = cur_page - 1; }
 else { show_page = cur_page + 1; }
userPage = prompt( msg, show_page );
if ( userPage > 0  ) {
	if ( userPage < 1 )     {    userPage = 1;  }
	if ( userPage > pages ) { userPage = pages; }
	if ( userPage == 1 )    {     start = 0;    }
	else { start = (userPage - 1) * per_page; }
	window.location = url_bit + "&st=" + start;
}
}
//-->
</script>
  <tbody><tr>
    <td style="border-bottom: 1px solid rgb(0, 0, 0);" align="left" height="60" valign="top" width="770">
	  <table border="0" cellpadding="0" cellspacing="0" height="60" width="770">
	    <tbody><tr>
		  <td height="43"><img src="style_images/Minimali-547/top1.jpg" border="0" height="43" width="770"></td>
		</tr>
		<tr>
		  <td class="toplinks" align="right" background="style_images/Minimali-547/top2.jpg" height="17">
		  &nbsp;[&nbsp;<a href="http://www.stealthbot.net/forumn/index.php?">Board Home</a>&nbsp;]&nbsp;&nbsp;&nbsp;[&nbsp;<a href="http://www.stealthbot.net/forumn/index.php?act=home">The New StealthBot.net</a>&nbsp;]&nbsp;&nbsp;&nbsp;[&nbsp;<a href="http://www.stealthbot.net/forum/getsb.php" target="_blank">Download</a>&nbsp;]&nbsp;&nbsp;&nbsp;[&nbsp;<a href="http://www.stealthbot.net/forumn/index.php?act=boardrules">Board Guidelines</a>&nbsp;]&nbsp;&nbsp;&nbsp;[&nbsp;<a href="http://www.stealthbot.net/forumn/index.php?act=Help">Help</a>&nbsp;]&nbsp;&nbsp;&nbsp;[&nbsp;<a href="http://www.stealthbot.net/forumn/index.php?act=Search&amp;f=">Search</a>&nbsp;]&nbsp;&nbsp;&nbsp;[&nbsp;<a href="http://www.stealthbot.net/forumn/index.php?act=Members">Members</a>&nbsp;]&nbsp;&nbsp;&nbsp;[&nbsp;<a href="http://www.stealthbot.net/forumn/index.php?act=calendar">Calendar</a>&nbsp;]<!--IBF.CHATLINK--><!--IBF.TSLLINK-->&nbsp;</td>
		</tr>
	  </tbody></table>
    </td>
  </tr>

  <tr>

    <td style="padding: 4px 10px 10px;" align="left" valign="top"> 
    
<div id="navstrip" align="left"><img src="style_images/Minimali-547/nav.gif" alt="&gt;" border="0">&nbsp;<a href="http://www.stealthbot.net/forumn/index.php?act=idx">StealthBot.net</a><img src="style_images/Minimali-547/navsep.gif" alt="&gt;" border="0">Download StealthBot</div>
<br> 
    <!--- AdSense block --->
    <center>
    <script type="text/javascript"><!--
        google_ad_client = "pub-6794476549115532";
        google_ad_width = 468;
        google_ad_height = 60;
        google_ad_format = "468x60_as";
        google_ad_type = "text_image";
        google_ad_channel ="";
        google_color_border = "204873";
        google_color_bg = "FFFFFF";
        google_color_link = "0000FF";
        google_color_text = "000000";
        google_color_url = "204873";
        //--></script>
        <script type="text/javascript"
          src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
    </script>
    </center>
    <br />
    <!--- End AdSense block --->

<div class="tableborder">
 <div class="pformstrip">Download StealthBot</div>
 <div class="tablepad">
 
 <?php randomMirror(); ?>
 
</div> 

 

<!-- Copyright Information -->
<br />
<div class="copyright" align="center">Powered by <a href="http://www.invisionboard.com/" target="_blank">Invision Power Board</a>(U) v1.3.1 Final © 2003 &nbsp;<a href="http://www.invisionpower.com/" target="_blank">IPS, Inc.</a></div>
<div align="center"><font class="copyright">Minimalist 2.0 Skin © <a href="mailto:skins@calliopedesigns.net">Ava</a> for <a href="http://www.calliopedesigns.net/">Calliope Designs</a></font></div>
<div align="center"><font class="copyright">© 2006 StealthBot.net Web Staff. See the <a href="http://legal.stealthbot.net/web/">Website Legal Information</a> page for terms of use, privacy policy and detailed copyright information.</font></div>
<br />

</td></tr></tbody></table>

</body></html>