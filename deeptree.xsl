<?xml version="1.0"  encoding="GB2312"?>           
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/TR/WD-xsl">           
<xsl:template match="/">          
<xsl:for-each select="xml/TreeNode">           
<div class="Node">
<xsl:attribute name="uid"><xsl:value-of select="@id"/></xsl:attribute>
<nobr>
<img type="icon" align="absmiddle">
<xsl:attribute name="id">icon<xsl:value-of select="@id"/></xsl:attribute>
<xsl:attribute name="src">
<xsl:choose>
<xsl:when test="child[. &gt; 0]">images/collapsed.gif</xsl:when>
<xsl:otherwise>images/endnode.gif</xsl:otherwise></xsl:choose></xsl:attribute>
</img><span type="Node" style="margin-left:3px">
<xsl:attribute name="title">
<xsl:choose>
<xsl:when test="title[. != '']"><xsl:value-of select="title"/></xsl:when>
<xsl:otherwise><xsl:value-of select="NodeText"/></xsl:otherwise></xsl:choose>
</xsl:attribute>
<xsl:if test="NodeUrl[. != '']">
<xsl:attribute name="url"><xsl:value-of select="NodeUrl"/></xsl:attribute>
<xsl:attribute name="target">
<xsl:choose>
<xsl:when test="target[. != '']"><xsl:value-of select="target"/></xsl:when>
<xsl:otherwise>main</xsl:otherwise></xsl:choose>
</xsl:attribute>
</xsl:if>
<xsl:attribute name="id">Node<xsl:value-of select="@id"/></xsl:attribute>
<xsl:value-of select="NodeText"/>
</span></nobr>
</div>
<xsl:if test="child[. &gt; 0]">
<div style="padding-left:12;display:none">
<xsl:attribute name="id">Child<xsl:value-of select="@id"/></xsl:attribute>
<div class="Node"><nobr><img src="images/endnode.gif" align="absmiddle"/><span class="NodeLoad">正在载入数据...</span></nobr></div>
</div>
</xsl:if>
</xsl:for-each>
</xsl:template> 
</xsl:stylesheet>