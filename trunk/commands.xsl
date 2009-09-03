<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
    <xsl:output method="text" />
    <xsl:template match="/">VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommandGeneratorObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

        <xsl:for-each select="/commands/command[not(@owner)]">
Private Sub cmd_<xsl:value-of select="@name"/>(ByRef oCommandDoc As clsCommandDocObj)

    Dim oCommand As clsCommandObj, oParameter As clsCommandParamsObj, oRestriction As clsCommandRestrictionObj

    With oCommandDoc
        If Not .OpenCommand("<xsl:value-of select="@name"/>", vbNullString) Then
            
            Call .CreateCommand("<xsl:value-of select="@name"/>", vbNullString, False)
            
            <xsl:if test="count(aliases/alias) > 0">
            With .aliases<xsl:for-each select="aliases/alias">
                .Add "<xsl:value-of select="text()"/>"</xsl:for-each>
            End With
            </xsl:if>
            
            <xsl:if test="count(documentation/description) > 0">
                .Description = <xsl:call-template name="vbstring"><xsl:with-param name="text" select="documentation/description"/></xsl:call-template>
            </xsl:if>
            
            <xsl:if test="count(documentation/specialnotes) > 0">
                .SpecialNotes = <xsl:call-template name="vbstring"><xsl:with-param name="text" select="documentation/specialnotes"/></xsl:call-template>
            </xsl:if>
            
            <xsl:if test="count(access/rank) > 0">
                .RequiredRank = "<xsl:value-of select="access/rank"/>"
            </xsl:if>
            
            <xsl:if test="count(access/flags/flag) > 0">
                .RequiredFlags = "<xsl:for-each select="access/flags/flag"><xsl:value-of select="normalize-space(text())"/></xsl:for-each>"
            </xsl:if>
            
            <xsl:for-each select="arguments/argument">
            Set oParameter = oCommandDoc.NewParameter("<xsl:value-of select="normalize-space(@name)"/>", <xsl:call-template name="getOptional" />, "<xsl:value-of select="normalize-space(@type)"/>")
            With oParameter
                <xsl:if test="count(documentation/description) > 0">
                    .Description = <xsl:call-template name="vbstring"><xsl:with-param name="text" select="documentation/description"/></xsl:call-template>
                </xsl:if>
                
                <xsl:if test="count(documentation/specialnotes) > 0">
                    .SpecialNotes = <xsl:call-template name="vbstring"><xsl:with-param name="text" select="documentation/specialnotes"/></xsl:call-template>
                </xsl:if>

                <xsl:if test="count(match) > 0">
                    .MatchMessage = <xsl:call-template name="vbstring"><xsl:with-param name="text" select="match/@message"/></xsl:call-template>
                    .MatchCaseSensitive = <xsl:call-template name="getMatchCaseSenitive" />
                    <xsl:if test="count(error) > 0">
                        .MatchError = <xsl:call-template name="vbstring"><xsl:with-param name="text" select="error/text()"/></xsl:call-template>
                    </xsl:if>
                </xsl:if>
                
                <xsl:for-each select="restrictions/restriction">
                    Set oRestriction = oCommandDoc.NewRestriction("<xsl:value-of select="normalize-space(@name)"/>", <xsl:call-template name="getRank" />, <xsl:call-template name="getFlags" />)
                    With oRestriction
                    <xsl:if test="count(access/rank) > 0">
                        .RequiredRank = "<xsl:value-of select="access/rank"/>"
                    </xsl:if>
                    <xsl:if test="count(access/flags/flag) > 0">
                        .RequiredFlags = "<xsl:for-each select="access/flags/flag"><xsl:value-of select="normalize-space(text())"/></xsl:for-each>"
                    </xsl:if>

                    <xsl:if test="count(match) > 0">
                        .MatchMessage = <xsl:call-template name="vbstring"><xsl:with-param name="text" select="match/@message"/></xsl:call-template>
                        .MatchCaseSensitive = <xsl:call-template name="getMatchCaseSenitive" />
                        <xsl:if test="count(error) > 0">
                            .MatchError = <xsl:call-template name="vbstring"><xsl:with-param name="text" select="error/text()"/></xsl:call-template>
                        </xsl:if>
                    </xsl:if>
                    End With
                </xsl:for-each>
            End With
            </xsl:for-each>
        End If
    End With
End Sub      
        </xsl:for-each>

Public Sub GenerateCommands()

    Dim oCommandDoc As New clsCommandDocObj
    <xsl:for-each select="/commands/command[not(@owner)]">
    Call cmd_<xsl:value-of select="@name"/>(oCommandDoc)</xsl:for-each>
	
	Call oCommandDoc.Save()
	
End Sub
        
    </xsl:template>

    <xsl:template name="getOptional">
        <xsl:choose>
            <xsl:when test="not(@optional)">True</xsl:when>
            <xsl:otherwise><xsl:choose>
                <xsl:when test="@optional = 1">False</xsl:when>
                <xsl:otherwise>True</xsl:otherwise></xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>    

    <xsl:template name="getMatchCaseSenitive">
        <xsl:choose>
            <xsl:when test="not(match[@case-sensitive])">True</xsl:when>
            <xsl:otherwise><xsl:choose>
                <xsl:when test="match[@case-sensitive] = 1">True</xsl:when>
                <xsl:otherwise>False</xsl:otherwise></xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>  

    <xsl:template name="getRank">
        <xsl:choose>
            <xsl:when test="count(access/rank) = 0">-1</xsl:when>
            <xsl:otherwise><xsl:value-of select="access/rank"/></xsl:otherwise>
        </xsl:choose>
    </xsl:template>  

    <xsl:template name="getFlags">
        <xsl:choose>
            <xsl:when test="count(access/flags/flag) = 0">vbNullstring</xsl:when>
            <xsl:otherwise>"<xsl:for-each select="access/flags/flag"><xsl:value-of select="text()"/></xsl:for-each>"</xsl:otherwise>
        </xsl:choose>
    </xsl:template>  
    
    <xsl:template name="vbstring">
        <xsl:param name="text"/>"<xsl:call-template name="replace-string">
            <xsl:with-param name="text" select="normalize-space($text)"/>
            <xsl:with-param name="from">&quot;</xsl:with-param>
            <xsl:with-param name="to">&quot;&quot;</xsl:with-param>
        </xsl:call-template>"
    </xsl:template>
    
    
    <!-- this example replaces ' with \' -->
    <!--
                    

    <xsl:call-template name="replace-string">
        <xsl:with-param name="text" select="normalize-space(./@Message)"/>
        <xsl:with-param name="from">'</xsl:with-param>
        <xsl:with-param name="to">\'</xsl:with-param>
    </xsl:call-template>
    -->
    <xsl:template name="replace-string">
        <xsl:param name="text"/>
        <xsl:param name="from"/>
        <xsl:param name="to"/>

        <xsl:choose>
            <xsl:when test="contains($text, $from)">

                <xsl:variable name="before" select="substring-before($text, $from)"/>
                <xsl:variable name="after" select="substring-after($text, $from)"/>
                <xsl:variable name="prefix" select="concat($before, $to)"/>

                <xsl:value-of select="$before"/>
                <xsl:value-of select="$to"/>
                <xsl:call-template name="replace-string">
                    <xsl:with-param name="text" select="$after"/>
                    <xsl:with-param name="from" select="$from"/>
                    <xsl:with-param name="to" select="$to"/>
                </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$text"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
</xsl:stylesheet>
