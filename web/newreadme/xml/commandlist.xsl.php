<?
ob_start("ob_gzhandler");
define("__READMEFILE", true);
header('Content-type: application/xml; charset="utf-8"',true);
require_once("../functions.php");
echo '<?xml version="1.0" encoding="utf-8"?>' . "\r\n";
?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:template match="/">
<? startPage("Command List"); ?>
    <h2>StealthBot Commands</h2>
    <p>Here is the full list of all <xsl:value-of select="count(/commands/command)" /> StealthBot commands. You can use them inside the bot by typing a slash, then the command:</p>
<p><b>/ban anAnnoyingPerson</b></p>
<p>You can use them inside the bot, but have their results displayed publicly to the channel you're in, by typing two slashes before the command:</p>
<p><b>//mp3<br />&lt;StealthBot&gt; Current MP3: 1418: Lamb of God - Again We Rise</b></p>

<p>You can also use them from outside the bot, if you have enough access:</p>
<p><b>&lt;Stealth&gt; .say Hello, world!<br />&lt;StealthBot&gt; Hello, world!</b></p>
<p>If you want to chain together multiple commands, you can do that also:</p>
<p><b>&lt;Stealth&gt; .whoami; say Hello, world!<br />&lt;StealthBot&gt; You have access 999 and flags ADMOST.<br />&lt;StealthBot&gt; Hello, world!</b></p>

<p>Lastly, if more than one of your bots share the same trigger, you can order them around by name:</p>
<p><b>&lt;Stealth&gt; .StealthBot say Hello, world!<br />&lt;StealthBot&gt; Hello, world!</b></p>
<p>Click on a command's name to see a detailed description of it, and an example of its use.</p>


<table width='600' cellpadding='1' cellspacing='15'>
    <tr valign='top'>
        <!-- 20 -->
        <td>
          <table cellpadding='4' cellspacing='2' style='border: 1px solid #000000'>
            <caption>Access level 20</caption>
            <tr class='headerrow'>
                <td>Command Word</td><td>Access</td><td>Alias(es)</td>
            </tr>
            <xsl:for-each select="/commands/command[access/rank = 20]">
              <xsl:sort select="@name"/>
              <tr>
                <xsl:if test="(position() mod 2 = 1)">
                  <xsl:attribute name="bgcolor">#CCCCCC</xsl:attribute>
                </xsl:if>
                <xsl:if test="(position() mod 2 = 0)">
                  <xsl:attribute name="bgcolor">#FFFFFF</xsl:attribute>
                </xsl:if>
                <td><a><xsl:attribute name="href">commands.php?xslt=commanddetails&amp;cmd=<xsl:value-of select="@name"/></xsl:attribute><xsl:value-of select="@name"/></a></td>
                <td><xsl:value-of select="access/rank"/></td>
                <td>
                    <xsl:for-each select="aliases/alias">
                        <xsl:value-of select="."/>
                    </xsl:for-each><br />
                </td>
              </tr>
            </xsl:for-each>
          </table>
        </td>
        <!-- 40 -->
        <td>
          <table cellpadding='4' cellspacing='2' style='border: 1px solid #000000'>
            <caption>Access level 40</caption>
            <tr class='headerrow'>
                <td>Command Word</td><td>Access</td><td>Alias(es)</td>
            </tr>
            <xsl:for-each select="/commands/command[access/rank = 40]">
              <xsl:sort select="@name"/>
              <tr>
                <xsl:if test="(position() mod 2 = 1)">
                  <xsl:attribute name="bgcolor">#CCCCCC</xsl:attribute>
                </xsl:if>
                <xsl:if test="(position() mod 2 = 0)">
                  <xsl:attribute name="bgcolor">#FFFFFF</xsl:attribute>
                </xsl:if>
                <td><a><xsl:attribute name="href">commands.php?xslt=commanddetails&amp;cmd=<xsl:value-of select="@name"/></xsl:attribute><xsl:value-of select="@name"/></a></td>
                <td><xsl:value-of select="access/rank"/></td>
                <td>
                    <xsl:for-each select="aliases/alias">
                        <xsl:value-of select="."/>
                    </xsl:for-each><br />
                </td>
              </tr>
            </xsl:for-each>
          </table>
        </td>      
      </tr>
    <tr valign='top'>
        <!-- 50 -->
        <td>
          <table cellpadding='4' cellspacing='2' style='border: 1px solid #000000'>
            <caption>Access level 50</caption>
            <tr class='headerrow'>
                <td>Command Word</td><td>Access</td><td>Alias(es)</td>
            </tr>
            <xsl:for-each select="/commands/command[access/rank = 50]">
              <xsl:sort select="@name"/>
              <tr>
                <xsl:if test="(position() mod 2 = 1)">
                  <xsl:attribute name="bgcolor">#CCCCCC</xsl:attribute>
                </xsl:if>
                <xsl:if test="(position() mod 2 = 0)">
                  <xsl:attribute name="bgcolor">#FFFFFF</xsl:attribute>
                </xsl:if>
                <td><a><xsl:attribute name="href">commands.php?xslt=commanddetails&amp;cmd=<xsl:value-of select="@name"/></xsl:attribute><xsl:value-of select="@name"/></a></td>
                <td><xsl:value-of select="access/rank"/></td>
                <td>
                    <xsl:for-each select="aliases/alias">
                        <xsl:value-of select="."/>
                    </xsl:for-each><br />
                </td>
              </tr>
            </xsl:for-each>
          </table>
        </td>
        <!-- 60 -->
        <td>
          <table cellpadding='4' cellspacing='2' style='border: 1px solid #000000'>
            <caption>Access level 60</caption>
            <tr class='headerrow'>
                <td>Command Word</td><td>Access</td><td>Alias(es)</td>
            </tr>
            <xsl:for-each select="/commands/command[access/rank = 60]">
              <xsl:sort select="@name"/>
              <tr>
                <xsl:if test="(position() mod 2 = 1)">
                  <xsl:attribute name="bgcolor">#CCCCCC</xsl:attribute>
                </xsl:if>
                <xsl:if test="(position() mod 2 = 0)">
                  <xsl:attribute name="bgcolor">#FFFFFF</xsl:attribute>
                </xsl:if>
                <td><a><xsl:attribute name="href">commands.php?xslt=commanddetails&amp;cmd=<xsl:value-of select="@name"/></xsl:attribute><xsl:value-of select="@name"/></a></td>
                <td><xsl:value-of select="access/rank"/></td>
                <td>
                    <xsl:for-each select="aliases/alias">
                        <xsl:value-of select="."/>
                    </xsl:for-each><br />
                </td>
              </tr>
            </xsl:for-each>
          </table>
        </td>
    </tr>
    <tr valign='top'>
        <!-- 70 -->
        <td>
          <table cellpadding='4' cellspacing='2' style='border: 1px solid #000000'>
            <caption>Access level 70</caption>
            <tr class='headerrow'>
                <td>Command Word</td><td>Access</td><td>Alias(es)</td>
            </tr>
            <xsl:for-each select="/commands/command[access/rank = 70]">
              <xsl:sort select="@name"/>
              <tr>
                <xsl:if test="(position() mod 2 = 1)">
                  <xsl:attribute name="bgcolor">#CCCCCC</xsl:attribute>
                </xsl:if>
                <xsl:if test="(position() mod 2 = 0)">
                  <xsl:attribute name="bgcolor">#FFFFFF</xsl:attribute>
                </xsl:if>
                <td><a><xsl:attribute name="href">commands.php?xslt=commanddetails&amp;cmd=<xsl:value-of select="@name"/></xsl:attribute><xsl:value-of select="@name"/></a></td>
                <td><xsl:value-of select="access/rank"/></td>
                <td>
                    <xsl:for-each select="aliases/alias">
                        <xsl:value-of select="."/>
                    </xsl:for-each><br />
                </td>
              </tr>
            </xsl:for-each>
          </table>
        </td>
        <!-- 80 -->
        <td>
          <table cellpadding='4' cellspacing='2' style='border: 1px solid #000000'>
            <caption>Access level 80</caption>
            <tr class='headerrow'>
                <td>Command Word</td><td>Access</td><td>Alias(es)</td>
            </tr>
            <xsl:for-each select="/commands/command[access/rank = 80]">
              <xsl:sort select="@name"/>
              <tr>
                <xsl:if test="(position() mod 2 = 1)">
                  <xsl:attribute name="bgcolor">#CCCCCC</xsl:attribute>
                </xsl:if>
                <xsl:if test="(position() mod 2 = 0)">
                  <xsl:attribute name="bgcolor">#FFFFFF</xsl:attribute>
                </xsl:if>
                <td><a><xsl:attribute name="href">commands.php?xslt=commanddetails&amp;cmd=<xsl:value-of select="@name"/></xsl:attribute><xsl:value-of select="@name"/></a></td>
                <td><xsl:value-of select="access/rank"/></td>
                <td>
                    <xsl:for-each select="aliases/alias">
                        <xsl:value-of select="."/>
                    </xsl:for-each><br />
                </td>
              </tr>
            </xsl:for-each>
          </table>
        </td>
    </tr>
    <tr valign='top'>
        <!-- 90 -->
        <td>
          <table cellpadding='4' cellspacing='2' style='border: 1px solid #000000'>
            <caption>Access level 90</caption>
            <tr class='headerrow'>
                <td>Command Word</td><td>Access</td><td>Alias(es)</td>
            </tr>
            <xsl:for-each select="/commands/command[access/rank = 90]">
              <xsl:sort select="@name"/>
              <tr>
                <xsl:if test="(position() mod 2 = 1)">
                  <xsl:attribute name="bgcolor">#CCCCCC</xsl:attribute>
                </xsl:if>
                <xsl:if test="(position() mod 2 = 0)">
                  <xsl:attribute name="bgcolor">#FFFFFF</xsl:attribute>
                </xsl:if>
                <td><a><xsl:attribute name="href">commands.php?xslt=commanddetails&amp;cmd=<xsl:value-of select="@name"/></xsl:attribute><xsl:value-of select="@name"/></a></td>
                <td><xsl:value-of select="access/rank"/></td>
                <td>
                    <xsl:for-each select="aliases/alias">
                        <xsl:value-of select="."/>
                    </xsl:for-each><br />
                </td>
              </tr>
            </xsl:for-each>
          </table>
        </td>
        <!-- 100 -->
        <td>
          <table cellpadding='4' cellspacing='2' style='border: 1px solid #000000'>
            <caption>Access level 100</caption>
            <tr class='headerrow'>
                <td>Command Word</td><td>Access</td><td>Alias(es)</td>
            </tr>
            <xsl:for-each select="/commands/command[access/rank = 100]">
              <xsl:sort select="@name"/>
              <tr>
                <xsl:if test="(position() mod 2 = 1)">
                  <xsl:attribute name="bgcolor">#CCCCCC</xsl:attribute>
                </xsl:if>
                <xsl:if test="(position() mod 2 = 0)">
                  <xsl:attribute name="bgcolor">#FFFFFF</xsl:attribute>
                </xsl:if>
                <td><a><xsl:attribute name="href">commands.php?xslt=commanddetails&amp;cmd=<xsl:value-of select="@name"/></xsl:attribute><xsl:value-of select="@name"/></a></td>
                <td><xsl:value-of select="access/rank"/></td>
                <td>
                    <xsl:for-each select="aliases/alias">
                        <xsl:value-of select="."/>
                    </xsl:for-each><br />
                </td>
              </tr>
            </xsl:for-each>
          </table>
        </td>
    </tr>    
</table>
<? require_once("../footer.php"); ?>
</xsl:template>
</xsl:stylesheet>