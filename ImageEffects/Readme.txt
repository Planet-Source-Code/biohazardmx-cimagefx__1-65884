cImageFX Readme
===============

  This class uses DIB's (Device Independent Bitmaps) to apply some entry-level effects to images
  in DC's. This version includes five effects:

	* Alpha Blend
	* Transparent Bit-Block-Transfer
	* Grayscale
	* Saturation
	* Color Fill

  Also are included three additional funcitons:

	* Screenshot (fullscreen or given area)
	* Wrapper to BitBlt
	* Wrapper to StretchBlt

  Image processing functions have support for mask color, so you can apply the effects only to
  the non-mask color pixels. Alpha blend it's faster than GetPixel-SetPixel operations because
  it uses DIB sections and generally it's fast.

  It's easy to use and it's free

Usage Instructions
==================
 
  To implement the class you'll need to add a reference by this code:

   Private WithEvents <ObjectName> As cImageFX
  
   Private Sub Form_Load()
     Set <ObjectName> = New cImageFX
   End Sub

  To use it you must have a source and a destination hDC. These could be two PictureBoxes or if
  you prefer you can use MemoryDC's, it's all the same:

   Private Sub Command1_Click()
     Call <ObjectName>.FxAlphaBlend(DestDC,DestX,DestY,DestHeight,DestWidth,SourceDC,...)
   End Sub

  If you use the WithEvents keyword you'll be able to receive progress events for the Image
  Processing procedures. Screenshot, BitBlt and StretchBlt doesn't raise progress events, you
  will get only the Complete event.

License
=======

  Please read the license in the cImageFX Class.

Contact
=======

  To contact the author:

    Go to http://biohazardmx.tripod.com/contact.htm and use the form to send me an email.