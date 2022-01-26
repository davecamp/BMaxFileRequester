SuperStrict
Framework Brl.StandardIO
Import Brl.SystemDefault 'RequestDir()
 
?Win32
 
Import "-lole32"
Import Pub.Win32
 
Global Shell32Dll:Byte Ptr = LoadLibraryA("Shell32.dll")
Global SHCreateItemFromParsingName:Int(pszPath$w,pbc:Byte Ptr,riid:Byte Ptr,ppv:IUnknown_ Var)"Win32" = GetProcAddress(Shell32Dll,"SHCreateItemFromParsingName")
 
Global CLSID_FileOpenDialog:Int[] = [$dc1c5a9c,$4ddee88a,$f860a1a5,$f7ae202a]
Global IID_IFileOpenDialog:Int[] = [$d57c7288,$4768d4ad,$969d02be,$60d93295]
Global IID_IShellItem:Int[] = [$43826d1e,$42eee718,$e2a155bc,$fe7bc361]
 
 
Extern"Win32"
        'These types are INCOMPLETE - DO NOT USE FOR ANYTHING ELSE !!!!!!
        Interface IModalWindow Extends IUnknown_
                Method Show(hWnd:Byte Ptr)
        EndInterface
 
        Interface IFileDialog Extends IModalWindow
                Method SetFileTypes()
                Method SetFileTypeIndex()
                Method GetFileTypeIndex()
                Method Advise()
                Method Unadvise()
                Method SetOptions:Int(dwOptions:Int)
                Method GetOptions:Int(dwOptions:Int Ptr)
                Method SetDefaultFolder:Int(pShellItem:IShellItem)
                Method SetFolder:Int(pSI:IShellItem)
                Method GetFolder()
                Method GetCurrentSelection()
                Method SetFilename:Int(pszName$w)
                Method GetFileName()
                Method SetTitle:Int(pszName$w)
                Method SetOKButtonLabel()
                Method SetFilenameLabel()
                Method GetResult:Int(pItem:IShellItem Var)
                Method AddPlace()
                Method SetDefaultExtension()
                Method Close()
                Method SetClientGuid()
                Method ClearClientData()
                Method SetFilter()
        EndInterface
       
        Interface IFileOpenDialog Extends IFileDialog
                Method GetResults:Int(ppEnum:IShellItemArray Ptr)
                Method GetSelectedItems:Int(ppsai:IShellItemArray Ptr)
        EndInterface
       
        Interface IShellItemArray Extends IUnknown_
                Method BindToHandler()
                Method GetPropertyStore()
                Method GetPropertyDescriptionList()
                Method GetAttributes()
                Method GetCount:Int(pdwNumItems:Int Ptr)
                Method GetItemAt:Int(dwIndex:Int, ppsi:IShellItem Ptr)
                Method EnumItems()
        EndInterface
       
        Interface IShellItem Extends IUnknown_
                Method BindToHandler()
                Method GetParent()
                Method GetDisplayName:Int(sigdnName:Int,ppszName:Short Ptr Var)
                Method GetAttributes()
                Method Compare()
        EndInterface
       
        Function CoCreateInstance:Int(rclsid:Byte Ptr,pUnkOuter:Byte Ptr,dwClsContext:Byte Ptr,riid:Byte Ptr,ppv:IUnknown_ Var)="HRESULT CoCreateInstance(REFCLSID, LPUNKNOWN, DWORD, REFIID, LPVOID)!"
        Function CoInitialize:Int(pvReserved:Byte Ptr)="HRESULT CoInitialize(LPVOID)!"
        Function CoUninitialize()="void CoUninitialize()!"
EndExtern
 
Function RequestFiles:String[](title:String, initialPath:String)
        Global pDialog:IFileOpenDialog
        Global pInitialPath:IShellItem
        Global pResults:IShellItemArray
        Local hr:Int
       
        CoInitialize(0)
 
        hr = CoCreateInstance(CLSID_FileOpenDialog,Null,CLSCTX_INPROC_SERVER,IID_IFileOpenDialog,pDialog)
        If hr < 0
		Cleanup()
		Return [RequestFile(Title,InitialPath)]
        EndIf
       
        Local dwOptions:Int
        pDialog.GetOptions(Varptr dwOptions)
        pDialog.SetOptions(dwOptions|$200) ' $200 = FOS_ALLOWMULTISELECT
 
        'Create an IShellItem for a default folder path
        InitialPath = InitialPath.Replace("/","\")
        If SHCreateItemFromParsingName(InitialPath,Null,IID_IShellItem,pInitialPath) < 0
		Cleanup()
		Return [RequestFile(Title,InitialPath)]
	EndIf
		
        If pDialog.SetFolder(pInitialPath) < 0
                CleanUp()
                Return [RequestFile(Title,InitialPath)]
        EndIf
 
        ' show it
        pDialog.SetTitle(Title)
        pDialog.Show(0)
 
        ' get the result
        If pDialog.GetResults(Varptr pResults) < 0
                CleanUp()
                Return Null
        EndIf
       
        'Get the results
        Local count:Int
        If pResults.GetCount(Varptr count) < 0
                CleanUp()
                Return Null
        EndIf
 
        Local selectedItemNames:String[count]
        For Local i:Int = 0 Until count
                Local pItem:IShellItem
                If pResults.getItemAt(i, Varptr pItem) >= 0
                        Local pName:Short Ptr
                        pItem.GetDisplayName($80058000,pName)
                        selectedItemNames[i] = String.FromWString(pName)
                EndIf
 
                If pItem pItem.Release_()
        Next 
       
        CleanUp()
        Return selectedItemNames
       
        Function CleanUp()
                If pDialog
                        pDialog.Release_()
                        pDialog = Null
                EndIf
                If pInitialPath
                        pInitialPath.Release_()
                        pInitialPath = Null
                EndIf
                If pResults
                        pResults.Release_()
                        pResults = Null
                EndIf
                CoUninitialize()
        EndFunction   
EndFunction
 
 
Function RequestFolder$(Title$,InitialPath$)
        Global pDialog:IFileOpenDialog
        Global pInitialPath:IShellItem
        Global pFolder:IShellItem
        Local hr:Int
        Local ResultFolder$
 
        CoInitialize(0)
 
        'Create Instance of the Dialog
        hr = CoCreateInstance(CLSID_FileOpenDialog,Null,CLSCTX_INPROC_SERVER,IID_IFileOpenDialog,pDialog)
 
        'Not on Vista or Win7?
        If hr < 0
		CleanUp()
		Return RequestDir(Title,InitialPath)
	EndIf
       
        'Set it to Browse Folders
        Local dwOptions:Int
        pDialog.GetOptions(Varptr dwOptions)
        pDialog.SetOptions(dwOptions|$20)
       
        'Set Title
        pDialog.SetTitle(Title)

        'Create an IShellItem for a default folder path
        InitialPath = InitialPath.Replace("/","\")
        hr = SHCreateItemFromParsingName(InitialPath, Null, IID_IShellItem, pInitialPath)
       
        If pDialog.SetFolder(pInitialPath) < 0
                CleanUp()
                Return RequestDir(Title,InitialPath)
        EndIf
               
        'Show the Dialog
        pDialog.Show(0)
 
        'Test the result
        If pDialog.GetResult(pFolder) < 0
                CleanUp
                Return ""
        EndIf
       
        'Get the result
        Local pName:Short Ptr
        pFolder.GetDisplayName($80058000,pName)
        ResultFolder = String.FromWString(pName)
       
        CleanUp()
        Return ResultFolder
       
        Function CleanUp()
                If pDialog
                        pDialog.Release_()
                        pDialog = Null
                EndIf
                If pInitialPath
                        pInitialPath.Release_()
                        pInitialPath = Null
                EndIf
                If pFolder
                        pFolder.Release_()
                        pFolder = Null
                EndIf
                CoUninitialize()
        EndFunction
EndFunction
?
