<div align="center">

## Tiled backgrounds on an MDI parent form? Yep\!


</div>

### Description



This code, which was inspired by a similar snippet of code by Ian Ippolito, permits tiling an image onto an MDI parent form's background. Getting an image onto an MDI parent is easy. Getting a tiled one is another story. We could try using a Clipboard operation, or build a big tiled background and save it and then laod it into the MDI parent's Picture property, but these are nasty, anal-retentive, and likely to simply not work. This code, however, works...
 
### More Info
 


Three inputs are required in the actual sub call:

MDIForm:    The name of the MDI parent form to tile an image onto.

bkgdtiler:   The name of a form that is used to generate the tiled background. Check the rest of the docs for the -required- parameters.

bkgdfile: The image file to load, with complete path.

The best place to put the call I find is in the form's Form_Load() event.



To use this code, you'll need to create a conventional (not an MDI parent or child) form and place PictureBox control on the form. Set the PictureBox's Name to Picture1, AutoRedraw, AutoSize, and ClipControls properties to TRUE, and its Visible property to FALSE.

The form also has certain requirements. Set its AutoRedraw and ClipControls properties to TRUE, its ControlBox and Visible properties to FALSE, remove its Caption, and set its BorderStyle property to 0 - None. When you call the sub, you'll be passing the form's name to the sub. The sub loads the form, tiles it, transfers the results to the MDI parent, and unloads the form.

DO NOT use this form for anything else as it's not kept in memory.



Nothing returned.



Larger forms (read: INSANELY large forms) might take some time to refresh on slow systems. Also, very small images take noticeably longer to tile than larger ones, so aim for about 125x125 pixel sizes for the tilable images.

Finally, although this code can use GIFs, interlaced GIFs tend to produce a single-pixel horizontal banding effect. So convert 'em to non-interlaced, etc.

Also, this does NOT work on forms other than MDI parent forms. I've submitted the non-MDI version of this code to this same location and this is the one for non-MDI-parent forms.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tom Honaker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tom-honaker.md)
**Level**          |Unknown
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tom-honaker-tiled-backgrounds-on-an-mdi-parent-form-yep__1-894/archive/master.zip)





### Source Code

```

Sub TileMDIBkgd(MDIForm As Form, bkgdtiler As Form, bkgdfile As String)
  If bkgdfile = "" Then Exit Sub
  Dim ScWidth%, ScHeight%
  ScWidth% = Screen.Width / Screen.TwipsPerPixelX
  ScHeight% = Screen.Height / Screen.TwipsPerPixelY
  Load bkgdtiler
  bkgdtiler.Height = Screen.Height
  bkgdtiler.Width = Screen.Width
  bkgdtiler.ScaleMode = 3
  bkgdtiler!Picture1.Top = 0
  bkgdtiler!Picture1.Left = 0
  bkgdtiler!Picture1.Picture = LoadPicture(bkgdfile)
  bkgdtiler!Picture1.ScaleMode = 3
  For n% = 0 To ScHeight% Step bkgdtiler!Picture1.ScaleHeight
    For o% = 0 To ScWidth% Step bkgdtiler!Picture1.ScaleWidth
      bkgdtiler.PaintPicture bkgdtiler!Picture1.Picture, o%, n%
    Next o%
  Next n%
  MDIForm.Picture = bkgdtiler.Image
  Unload bkgdtiler
End Sub
```

