<?
ob_start("ob_gzhandler");
define("__READMEFILE", true);
header('Content-type: application/xml; charset="utf-8"',true);
require_once("../functions.php");
// get the command from the query string
$cmd = $_GET['cmd'];
echo '<?xml version="1.0" encoding="utf-8"?>' . "\r\n";
?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:template match="/">
<xsl:for-each select="commands/command[@name='<?=$cmd;?>']">
<?
startPage("Command: " . $cmd);
?>
    <h2><xsl:value-of select="@name" /> (<xsl:value-of select="access/rank" />)</h2>
    <p><b>Alias<xsl:if test="count(aliases/alias) &gt; 1">es</xsl:if></b><br />
    <xsl:if test="count(aliases/alias) &gt; 0">
        <xsl:for-each select="aliases/alias">
            <xsl:value-of select="." /> 
            <xsl:if test="position()!=last()">
                <xsl:text>, </xsl:text>
           </xsl:if>
        </xsl:for-each>
     </xsl:if>
     <xsl:if test="count(aliases/alias) = 0">
        <xsl:text>None</xsl:text>
     </xsl:if>
    </p>
    <p><b>Description</b><br /><xsl:value-of select="documentation/description" /></p>
    <p><b>Argument<xsl:if test="count(arguments/argument) &gt; 1">es</xsl:if></b><br />
    <xsl:if test="count(arguments/argument) &gt; 0">
        <ul>
            <xsl:for-each select="arguments/argument">
                <li><b><xsl:value-of select="@name" /></b>
                <xsl:if test="@optional = 'true'">
                    <xsl:text> (optional)</xsl:text>
                </xsl:if>
                <br /><xsl:value-of select="documentation/description" />
                </li>
            </xsl:for-each>
        </ul>
    </xsl:if>
    <xsl:if test="count(arguments/argument) = 0">
        <xsl:text>No Arguments</xsl:text>
    </xsl:if>
    </p>
    <p><a href='commands.php?xslt=commandlist'>Back to the list of commands?</a></p>
    
<? require_once("../footer.php"); ?>
</xsl:for-each>
</xsl:template>
</xsl:stylesheet>