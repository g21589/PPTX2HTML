PPTX2HTML - Developer Notes
==========

Pixel = EMUs * Resolution / 914400;
where "Resolution" is resolution of your screen. e.g. 96 dpi.

<p:cSld> common slide data element
  <p:sp> shapes element
    <p:nvSpPr> Each shape element contains a set of non-visual properties
      <p:nvPr> non-visual properties
	    <p:ph> a placeholder
		  The placeholder element is empty but does have several possible attributes.
		  It is using the "idx" and "type" attributes that the shapes are linked across the three slide types.
