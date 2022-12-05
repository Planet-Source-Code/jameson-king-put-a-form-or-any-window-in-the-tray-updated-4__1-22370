﻿<div align="center">

## Put a form \(or any window\) in the TRAY\!  \[Updated \#4\]

<img src="PIC20014221659258835.jpg">
</div>

### Description

In my endeavor the make a weather program for my computer that would scroll weather information, I came to make this code. I was looking for a place to scroll the information without bothering any other programs and without the form taking focus. I was looking for the perfect place. I found it but it was of course occupied by something else. That something else was the task tray (Those small little icons and the clock in the bottom right hand corner of most windows computers). Well I said to myself how shall I take that area. I thought I would simply shape my form to that of the trays window. Then make my form a child of that window and set it's position to the topmost. Well my idea was correct and it works perfectly. You can use this technique for any window, it is far less involved then many other attempts witch require sub-classing. The window will resize to the RECT of the tray when the Task bar is resize and or moved and call the form to repaint also. I hope all of you can get something useful out of this I know have.

*NEW* YOU can now resize the tray by doing a call ResizeTray (Number to resize by,A 1 or a 2 for rither making it larger or smaller)

so

call resizetray 11,1

would add eleven icons to the tray

and

call resizetray 11,2

would subtract eleven icons
 
### More Info
 
You may use this code freely and for whatever purpose you wish, I ask only that you give me FULL credit for the code as is (Only the code in the WFU_Main Bas and the WFU_Easy_Sub bas.) and that you Include everything I have stated from my header at the top (which includes my full name the times I created this and the date) and also all text below the header up to and including this text you are reading now. In addition, I herby define code as being anything written inside the those basses EXCEPT the comments.


<span>             |<span>
---                |---
**Submitted On**   |2001-04-22 17:20:08
**By**             |[Jameson King](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jameson-king.md)
**Level**          |Advanced
**User Rating**    |5.0 (70 globes from 14 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Put a form186904222001\.zip](https://github.com/Planet-Source-Code/jameson-king-put-a-form-or-any-window-in-the-tray-updated-4__1-22370/archive/master.zip)








