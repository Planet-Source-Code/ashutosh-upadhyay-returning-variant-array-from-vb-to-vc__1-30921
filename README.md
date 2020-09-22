<div align="center">

## returning variant array from  vb to vc\+\+

<img src="Image3.gif">
</div>

### Description

Passing variant array between vb server and vc++ is noteasy task. I found a technique when i searched for a week on this matarial - how to return variant array from activeX exe(VB) to MFC(VC++).To make easy you , i submit not in exact form what i found on msdn and internet ,but what technique i used and got success.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |1998-01-01 06:22:50
**By**             |[Ashutosh Upadhyay](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ashutosh-upadhyay.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[returning\_495361182002\.zip](https://github.com/Planet-Source-Code/ashutosh-upadhyay-returning-variant-array-from-vb-to-vc__1-30921/archive/master.zip)





### Source Code

```
<h4>How to return variant array from VB ActiveX Server (EXE/DLL) to VC++</h4><p>
<pre>
Step 1. Make a VB ActiveXDll /EXE project<p>
   1.1 - Name your Project and Default Class , We assume Prject name - Project1 and      Class Name - Class1<p>
   1.2 copy and past follwoing function
     public function abc() as Variant
     dim a(10) as long
     a(0) = 10
     a(1) = 13
     .
     .
     .
     a(10) = 67
    end function
   1.3  open file menu ,choose make project1.dll/project1.exe
Step 2. Open A Simple VC++ Project
   2.1 choose MFC application(exe),then,dialog application,then,support automation
   2.2 choose View -> ClassWizard -> Automation tab -> Press Add Class ->     choose From a type library
   2.3 allow MFC to generate wrapper classes for your VB server. We assume name     of Wrapper class genrated by MFC is _Class1
   2.4 in your dialog class ( in ..dlg.h file) , make a variable of your wrapper class
      class YourApplictionDlg : public CDialog
      {
      ..........
      ............
      public :
      _Class1 ashu;
      ...............
     ................
      ............
      };
     2.5 open Oninitdialog() function of your dialog (dlg) class and copy following       code in the end of function
      // TODO: Add extra initialization here
      ashu.CreateDispatch("Project1.Class1");
      return TRUE;
    2.6 To retrieve the values of array of VB server write followoing code
       in a function, for example in OnOK()
void .......Dlg::OnOK() {
 VARIANT v;
 v = ashu.abc();
 SAFEARRAY *parray;
 parray = v.parray;
short sElem;
long lLb, lUb, l;
long lResult[10];
if (parray == NULL){ // array has not been initialized
MessageBox("NULL");
return;
}
if ((parray)->cDims != 1) {// check number of dimensions
 MessageBox("Dim <> 1");
TRACE("%d\n",parray->cDims);
return ;
}
TRACE("%d\n",parray->cDims);
// get the upper and lower bounds of the array
if (FAILED(SafeArrayGetLBound(parray, 1, &lLb)) ||
	FAILED(SafeArrayGetUBound(parray, 1, &lUb))){
 MessageBox("Array Bound Failed");
return ;
}
TRACE(" %d %d\n",lLb,lUb);
// loop through the array and put the elements into array lResult
int i=0;
for (l = lLb; l <= lUb; l++) {
	if (FAILED(SafeArrayGetElement(parray, &l, &sElem))){
MessageBox("Element failed");
return ; }
TRACE("%d \n",sElem);
lResult[i++] = sElem;
}
 ///// You can display array in message box
///// To display
 CString s;
 s.Format("%d %d %d .....  %d", lResult[0],lResult[1],..........,lResult[lUb]);
Messagebox(s);
///////////// delete CDialog::OnOk()
}
/////////////////
    2.7 Now in your dialog class OnClose or OnCancel function copy following
       ashu.DetachDispatch();
	  ashu.ReleaseDispatch();
    2.8 build your program and run. you may add <afxole.h> header file in your      _Class1.h file, if any problem in building.
 </pre>
```

