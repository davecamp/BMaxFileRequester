Strict
 
?Win32
 
Import "-lole32"
Import Pub.Win32
 
Global Shell32Dll = LoadLibraryA("Shell32.dll")
Global SHCreateItemFromParsingName(pszPath$w,pbc:Byte Ptr,riid:Byte Ptr,ppv:IUnknown Var)"Win32" = GetProcAddress(Shell32Dll,"SHCreateItemFromParsingName")
 
Global CLSID_FileOpenDialog[] = [$dc1c5a9c,$4ddee88a,$f860a1a5,$f7ae202a]
Global IID_IFileOpenDialog[] = [$d57c7288,$4768d4ad,$969d02be,$60d93295]
Global IID_IShellItem[] = [$43826d1e,$42eee718,$e2a155bc,$fe7bc361]
 
 
Extern"Win32"
        'These types are INCOMPLETE - DO NOT USE FOR ANYTHING ELSE !!!!!!
        Type IModalWindow Extends IUnknown
                Method Show(hWnd)
        EndType
 
        Type IFileDialog Extends IModalWindow
                Method SetFileTypes()
                Method SetFileTypeIndex()
                Method GetFileTypeIndex()
                Method Advise()
                Method Unadvise()
                Method SetOptions(dwOptions)
                Method GetOptions(dwOptions Ptr)
                Method SetDefaultFolder(pShellItem:Byte Ptr)
                Method SetFolder(pSI:Byte Ptr)
                Method GetFolder()
                Method GetCurrentSelection()
                Method SetFilename(pszName$w)
                Method GetFileName()
                Method SetTitle(pszName$w)
                Method SetOKButtonLabel()
                Method SetFilenameLabel()
                Method GetResult(pItem:IShellItem Var)
                Method AddPlace()
                Method SetDefaultExtension()
                Method Close()
                Method SetClientGuid()
                Method ClearClientData()
                Method SetFilter()
        EndType
       
        Type IFileOpenDialog Extends IFileDialog
                Method GetResults(ppEnum:IShellItemArray Ptr)
                Method GetSelectedItems(ppsai:IShellItemArray Ptr)
        EndType
       
        Type IShellItemArray Extends IUnknown
                Method BindToHandler()
                Method GetPropertyStore()
                Method GetPropertyDescriptionList()
                Method GetAttributes()
                Method GetCount(pdwNumItems:Int Ptr)
                Method GetItemAt(dwIndex:Int, ppsi:IShellItem Ptr)
                Method EnumItems()
        EndType
       
        Type IShellItem Extends IUnknown
                Method BindToHandler()
                Method GetParent()
                Method GetDisplayName(sigdnName,ppszName:Short Ptr Var)
                Method GetAttributes()
                Method Compare()
        EndType
       
        Function CoCreateInstance(rclsid:Byte Ptr,pUnkOuter:Byte Ptr,dwClsContext,riid:Byte Ptr,ppv:IUnknown Var) 'My version
        Function CoInitialize(pvReserved)
        Function CoUninitialize()
EndExtern
 
Function RequestFiles:String[](title:String, initialPath:String)
        Global pDialog:IFileOpenDialog
        Global pInitialPath:IShellItem
        Global pResults:IShellItemArray
        Local hr:Int
       
        CoInitialize(0)
 
        hr = CoCreateInstance(CLSID_FileOpenDialog,Null,CLSCTX_INPROC_SERVER,IID_IFileOpenDialog,pDialog)
        If hr < 0
                CleanUp()
                Return Null
        EndIf
       
        Local dwOptions:Int
        pDialog.GetOptions(Varptr dwOptions)
        pDialog.SetOptions(dwOptions|$200) ' $200 = FOS_ALLOWMULTISELECT
 
        'Create an IShellItem for a default folder path
        InitialPath = Replace(InitialPath,"/","")
        SHCreateItemFromParsingName(InitialPath,Null,IID_IShellItem,pInitialPath)
       
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
        Local hr,ResultFolder$
 
        CoInitialize(0)
 
        'Create Instance of the Dialog 
        hr = CoCreateInstance(CLSID_FileOpenDialog,Null,CLSCTX_INPROC_SERVER,IID_IFileOpenDialog,pDialog)
 
        'Not on Vista or Win7?
        If hr < 0 CleanUp(); Return RequestDir(Title,InitialPath)
       
        'Set it to Browse Folders
        Local dwOptions
        pDialog.GetOptions(Varptr dwOptions)
        pDialog.SetOptions(dwOptions|$20)
       
        'Set Title
        pDialog.SetTitle(Title)
       
        'Create an IShellItem for a default folder path
        InitialPath = Replace(InitialPath,"/","")
        SHCreateItemFromParsingName(InitialPath,Null,IID_IShellItem,pInitialPath)
       
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
