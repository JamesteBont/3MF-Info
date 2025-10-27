;3MF-Info 1.0.0.0
;Orca-FlashForge metatag editor
;26-10-2025
;Theo van Duin

EnableExplicit
UseZipPacker()

Global Window_0, StatusBar_0
Global XMLModel, ListIcon_0, ModelFile$, OrcaFile.s
Global row.l, old.s, new.s, infotext.s
Global startdir.s, UserDir.s, fcount.l

startdir = GetCurrentDirectory()
UserDir = GetUserDirectory(#PB_Directory_Downloads) ;we paken alles uit in de Download systeemmap

infotext = "3MF-Info 1.0"+Chr(13)
infotext = infotext+ ""+Chr(13)
infotext = infotext+ "Orca-FlashForge prevents users from editing items in the Project tab."+Chr(13)
infotext = infotext+ "This program extracts the selected .3MF files to the Downloads folder."+Chr(13)
infotext = infotext+ "And allows you To edit the metatags in the /3D/3dmodel.model file."+Chr(13)
infotext = infotext+ "When saved, all files are compressed "+Chr(13)
infotext = infotext+ "back into a single .3MF file."+Chr(13)
infotext = infotext+ ""+Chr(13)
infotext = infotext + "Use at your own risk. Software is developed For our own use, but may be Shared."+Chr(13)
infotext = infotext+ ""+Chr(13)
infotext = infotext+ "Developed With PureBasic 6.11 LTS (Windows - x64)"+Chr(13)
infotext = infotext+ "Tested With Orca Flashforge 1.4.2"+Chr(13)
infotext = infotext+ "Theo van Duin 26-10-2025"

Enumeration FormGadget
  #String_3
EndEnumeration

Enumeration FormMenu
  #MenuItem_2
  #MenuItem_3
  #MenuItem_4
  #MenuItem_5
  #MenuItem_7
EndEnumeration

Procedure OpenWindow_0()
  Window_0 = OpenWindow(#PB_Any, 50, 50, 950, 400, "3MF-info", #PB_Window_SystemMenu|#PB_Window_MinimizeGadget)
    If CreateStatusBar(StatusBar_0, WindowID(Window_0))
    AddStatusBarField(800)
    AddStatusBarField(#PB_Ignore)
  EndIf
  CreateMenu(0, WindowID(Window_0))
  MenuTitle("File")
  MenuItem(#MenuItem_2, "Open...")
  MenuItem(#MenuItem_3, "Save")
  MenuItem(#MenuItem_4, "Save As...")
  MenuBar()
  MenuItem(#MenuItem_5, "Quit")
  MenuTitle("Help")
  MenuItem(#MenuItem_7, "About 3MF info text")
  ListIcon_0 = ListIconGadget(#PB_Any, 8, 8, 930, 344, "Metadata", 300)
  AddGadgetColumn(ListIcon_0, 1, "Value", 600)
EndProcedure


Procedure.s EditDescription(_txt.s)
  Protected Window_1, Editor_0, Button_0, Button_1, newtxt.s, Event, QuitEdit
  Window_1 = OpenWindow(#PB_Any, 10, 10, 624, 280, "Description", #PB_Window_SystemMenu|#PB_Window_ScreenCentered)
  Editor_0 = EditorGadget(#PB_Any, 8, 8, 608, 232)
  Button_0 = ButtonGadget(#PB_Any, 8, 248, 80, 24, "Ok")
  Button_1 = ButtonGadget(#PB_Any, 96, 248, 80, 24, "Cancel")
  SetGadgetText(Editor_0, _txt)
  
    Repeat
    Event = WaitWindowEvent() 
    Select Event
      Case  #PB_Event_CloseWindow
        newtxt = ""
        QuitEdit = 1
      Case  #PB_Event_Gadget
        Select EventGadget()
          Case Button_0 ;Ok        
            newtxt = GetGadgetText(Editor_0)
            QuitEdit = 1
          Case Button_1
            newtxt = ""
            QuitEdit = 1
        EndSelect
    EndSelect
    
  Until QuitEdit = 1
  
  CloseWindow(Window_1)
  ProcedureReturn newtxt
EndProcedure


Procedure.s EditLicense(_txt.s)
  Protected Window_2, Combo_0, Button_0, Button_1, newtxt.s, Event, QuitEdit,result$
  Window_2 = OpenWindow(#PB_Any, 10, 10, 624, 72, "License", #PB_Window_SystemMenu|#PB_Window_ScreenCentered)
  Combo_0 = ComboBoxGadget(#PB_Any, 8, 8, 608, 21)
  Button_0 = ButtonGadget(#PB_Any, 8, 40, 80, 24, "Ok")
  Button_1 = ButtonGadget(#PB_Any, 96, 40, 80, 24, "Cancel")
  AddGadgetItem(Combo_0,-1,"CC0 - Creative Commons Zero")
  AddGadgetItem(Combo_0,-1,"BY - Attribution")
  AddGadgetItem(Combo_0,-1,"BY-SA - Attribution ShareAlike")
  AddGadgetItem(Combo_0,-1,"BY-ND - Attribution NoDerivatives")
  AddGadgetItem(Combo_0,-1,"BY-NC - Attribution NonCommercial")
  AddGadgetItem(Combo_0,-1,"BY-NC-SA - Attribution NonCommercial ShareAlike")  
  AddGadgetItem(Combo_0,-1,"BY-NC-ND - Attribution NonCommercial NoDerivatives")

    Repeat
    Event = WaitWindowEvent() 
    Select Event
      Case  #PB_Event_CloseWindow
        newtxt = ""
        QuitEdit = 1
      Case  #PB_Event_Gadget
        Select EventGadget()
          Case Button_0 ;Ok    
            newtxt = GetGadgetText(Combo_0)
            Select newtxt
              Case "CC0 - Creative Commons Zero"
                newtxt = "CC0"
              Case "BY - Attribution" 
                newtxt = "BY"
              Case "BY-SA - Attribution ShareAlike" 
                newtxt = "BY-SA"                
              Case "BY-ND - Attribution NoDerivatives"
                newtxt = "BY-ND"
              Case "BY-NC - Attribution NonCommercial"
                newtxt = "BY-NC"                
              Case "BY-NC-SA - Attribution NonCommercial ShareAlike"
                newtxt = "BY-NC-SA"                
              Case "BY-NC-ND - Attribution NonCommercial NoDerivatives"
                newtxt = "BY-NC-ND"
            EndSelect
            QuitEdit = 1
          Case Button_1
            newtxt = ""
            QuitEdit = 1
        EndSelect
    EndSelect
    
  Until QuitEdit = 1
  CloseWindow(Window_2)
  ProcedureReturn newtxt
EndProcedure


Procedure.i DirectoryExists(cDir.s)
  Protected CurDir.s, rslt.i
    CurDir = GetCurrentDirectory()
    rslt = SetCurrentDirectory(cDir)
    SetCurrentDirectory(CurDir)
  ProcedureReturn rslt
EndProcedure 


 Procedure.i Unzip3MF()
  Protected filename.s, pattern.s, temppath.s, curdir.s, newname.s, result_text.s, infotext.s
  Protected rslt.l, pos.l, start.l, counter.l, keypressed.i
  SetCurrentDirectory(UserDir)
  
  If DirectoryExists(UserDir + "Temp_3mf_\")                         ;Tijdelijke map bestaat
    infotext = UserDir + "Temp_3mf_\ already exists!" + Chr(13) + "Do you want to delete this folder?"
    keypressed = MessageRequester("3MF info text",infotext, #PB_MessageRequester_Warning|#PB_MessageRequester_YesNo)
    If keypressed = #PB_MessageRequester_Yes
      rslt = DeleteDirectory(UserDir + "Temp_3mf_\","",#PB_FileSystem_Recursive|#PB_FileSystem_Force)  ;alles opruimen
    Else
      ProcedureReturn #False       
    EndIf
  EndIf

  pattern = "Orca slicer file (*.3MF)|*.3MF|All files (*.*)|*.*"
  filename = OpenFileRequester("Open Orca slicer file", "", pattern, 0 )
  If filename <>""                                                   ;er is een bestand gekozen 
    If OpenPack(0, filename)                                         ;open 3MF zip bestand
      StatusBarText(StatusBar_0,1,"OpenPack") 
      temppath = UserDir + "Temp_3mf_\"
      rslt = CreateDirectory(temppath)                               ;maak een tijdelijke map      
      If rslt                                                        ;Tijdelijke map ok
        If ExaminePack(0)                                            ;3MF onderzoeken
          While NextPackEntry(0)                                     ;Zolang er bestanden zijn
            newname = ReplaceString(PackEntryName(0),"/","\")        ;gebruik alleen backslash Compatibel met 3MF/RMF, ZIP, en andere OPC-formaten
            start = 1
            pos = 0
            While FindString(newname,"\",start)                      ;vind alle submappen 
              pos = FindString(newname,"\",start) 
              start = pos + 1
            Wend  
            
            If pos                                                   ;er zit een submap in de naam
              rslt = SetCurrentDirectory(temppath+Left(newname,pos)) ;probeer of de map al bestaat
              If rslt = 0                                            ;map bestaat niet
                rslt = CreateDirectory(temppath+Left(newname,pos))   ;maak hem dan
              EndIf
            EndIf 
            
            If UncompressPackFile(0, temppath+newname, PackEntryName(0)) = -1  ;als uitpakken niet lukt
              ClosePack(0)
              StatusBarText(StatusBar_0,0,"Unpacking error!")  
              StatusBarText(StatusBar_0,1,Str(counter))
              MessageRequester("Unpacking 3MF", "Couldn't unpack " + PackEntryName(0) + ".", #PB_MessageRequester_Error)
            Else                                                     ;Bestand is succesvol uitgepakt
              counter = counter + 1
              result_text = result_text + Str(counter) + ": " + temppath + newname + " extracted." + Chr(13)
            EndIf
          Wend                                                       ;alle bestanden uitgepakt
          ClosePack(0)
          StatusBarText(StatusBar_0,0,"Files successfully extracted.")  
          StatusBarText(StatusBar_0,1,Str(counter))
        Else                                                        ;kan 3MF niet onderzoeken
          StatusBarText(StatusBar_0,1,"ClosePack 1") 
          ClosePack(0)
          StatusBarText(StatusBar_0,0,"Cannot extract 3MF file!") 
          StatusBarText(StatusBar_0,1,"0")
          ProcedureReturn #False
        EndIf                                                        ;einde 3MF onderzoeken
      Else                                                           ;kon geen tijdelijke map maken
        ClosePack(0)
        StatusBarText(StatusBar_0,0,"Cannot create temporary folder!")
        StatusBarText(StatusBar_0,1,"0")
        ProcedureReturn #False
      EndIf                                                          ;einde tijdelijke map
    Else                                                             ;OpenPack() mislukt
      StatusBarText(StatusBar_0,0,"Cannot open this file!")
      StatusBarText(StatusBar_0,1,"0")
      ProcedureReturn #False
    EndIf                                                            ;einde Openpack()
    
  Else                                                               ;er is geen bestand gekozen
    ProcedureReturn #False
  EndIf                                                              ;einde bestandskeuze
  
  OrcaFile = filename
  SetWindowTitle(Window_0, "3MF-Info - "+ OrcaFile) 
  ProcedureReturn #True
EndProcedure


Procedure OpenModel(FileName$)
  Protected xmlRoot, xmlNode, child, name$, value$
  
  If XMLModel
    FreeXML(XMLModel)
  EndIf
  
  XMLModel = LoadXML(#PB_Any, FileName$)
  If XMLModel
    xmlRoot = RootXMLNode(XMLModel)
    If xmlRoot
      xmlNode = XMLNodeFromPath(xmlRoot, "/model")
      If xmlNode                                                     ;hier kun je de metadata uitlezen
        
      Else
        MessageRequester("Error", "No <model> node found.")
      EndIf
    Else
      MessageRequester("Error", "No root node in XML.")
    EndIf
  Else
    MessageRequester("Fout", "Could not load XML.")
  EndIf

  ClearGadgetItems(ListIcon_0)  
  child = ChildXMLNode(xmlNode)
  While child
    If XMLNodeType(child) = #PB_XML_Normal
      If LCase(GetXMLNodeName(child)) = "metadata"
        name$  = GetXMLAttribute(child, "name")
        value$ = GetXMLNodeText(child)
        AddGadgetItem(ListIcon_0, -1, name$ + Chr(10) + value$)
      EndIf
    EndIf
    child = NextXMLNode(child)
  Wend
  
  ProcedureReturn #True
EndProcedure


Procedure PackDirectoryRecursive(BasePath.s, SubPath.s)
  Protected currentPath.s, dirID, entryName.s, nextSubPath.s, sourceFile.s, targetFile.s, cnt.i

  If Right(BasePath, 1) <> "\"                                       ;Zorg dat BasePath altijd eindigt met een backslash
    BasePath + "\"
  EndIf

  currentPath = BasePath + SubPath
  dirID = ExamineDirectory(#PB_Any, currentPath, "*.*")
  If dirID
    While NextDirectoryEntry(dirID)
      entryName = DirectoryEntryName(dirID)

      If DirectoryEntryType(dirID) = #PB_DirectoryEntry_File
        sourceFile = currentPath + entryName
        targetFile = SubPath + entryName
        targetFile = ReplaceString(targetFile,"\","/")               ;gebruik alleen slash bij inpakken Compatibel met 3MF/RMF, ZIP, en andere OPC-formaten
        AddPackFile(0, sourceFile, targetFile)
        fcount +1
      ElseIf DirectoryEntryType(dirID) = #PB_DirectoryEntry_Directory
        If entryName <> "." And entryName <> ".."
          nextSubPath = SubPath + entryName + "\"
          nextSubPath = ReplaceString(nextSubPath,"\","/")           ;gebruik alleen slash bij inpakken Compatibel met 3MF/RMF, ZIP, en andere OPC-formaten
          PackDirectoryRecursive(BasePath, nextSubPath)
        EndIf
      EndIf
    Wend
    FinishDirectory(dirID)
  EndIf
EndProcedure


Procedure SaveModel(FileName$)
  Protected xmlRoot, xmlModelNode, child
  Protected i, name$, value$, found

  If XMLModel = 0
    MessageRequester("Error", "No model loaded.")
    ProcedureReturn #False
  EndIf

  xmlRoot = RootXMLNode(XMLModel)                                    ; Haal de root-node op van het document
  If xmlRoot = 0
    MessageRequester("Error", "No valid XML root.")
    ProcedureReturn #False
  EndIf

  xmlModelNode = XMLNodeFromPath(xmlRoot, "model")                   ; Zoek het <model> element onder de root
  If xmlModelNode = 0
    MessageRequester("Error", "Unable to find <model>.")
    ProcedureReturn #False
  EndIf

  For i = 0 To CountGadgetItems(ListIcon_0) - 1                      ; Loop door de items in de ListIcon
    name$  = GetGadgetItemText(ListIcon_0, i, 0)
    value$ = GetGadgetItemText(ListIcon_0, i, 1)
    found  = #False
    child = ChildXMLNode(xmlModelNode)
    While child
      If XMLNodeType(child) = #PB_XML_Normal
        If LCase(GetXMLNodeName(child)) = "metadata"
          If GetXMLAttribute(child, "name") = name$
            SetXMLNodeText(child, value$)
            found = #True
            Break
          EndIf
        EndIf
      EndIf
      child = NextXMLNode(child)
    Wend

    If Not found                                                     ; Voeg toe als het nog niet bestaat
      child = CreateXMLNode(xmlModelNode, "metadata", -1)
      SetXMLAttribute(child, "name", name$)
      SetXMLNodeText(child, value$)
    EndIf
  Next i

  If SaveXML(XMLModel, FileName$)
    MessageRequester("OK", "Metadata saved as " + FileName$)
  Else
    MessageRequester("Error", "Could not save XML file.")
  EndIf

  ProcedureReturn #True
EndProcedure



Procedure Save3MF(fname.s)
  Protected pattern.s, modelfile.s
  
  modelfile = UserDir + "Temp_3mf_\3D\3dmodel.model"
  If SaveModel(modelfile)
  pattern = "Orca slicer file (*.3MF)|*.3MF|All files (*.*)|*.*"
  
  If fname = ""
    fname = SaveFileRequester("Save As...", "", pattern, 0 )
  EndIf
  
  If fname
    If DirectoryExists(UserDir + "Temp_3mf_\")   
      If CreatePack(0, fname, #PB_PackerPlugin_Zip)
        fcount = 0
        PackDirectoryRecursive(UserDir + "Temp_3mf_\", "")
        ClosePack(0)
        MessageRequester("Finished", "Folder " + UserDir + "Temp_3mf_\" + " is completely compressed into " + fname + ".")
        StatusBarText(StatusBar_0,0,"Temp_3mf_\" + " is completely compressed into  " + fname ) 
        StatusBarText(StatusBar_0,1,Str(fcount))
      Else
        MessageRequester("Error", "Unable to create zip file: " + fname)
      EndIf
    EndIf
  EndIf
  EndIf
EndProcedure  


Procedure ChangeItem(selected_row)
  If selected_row < 0 Or selected_row >= CountGadgetItems(ListIcon_0)
    ;MessageRequester("Error", "Invalid row selected. "+ Str(selected_row))
    ProcedureReturn
  EndIf

  ; Haal de MetaTag van de geselecteerde rij op
  Protected tag$ = GetGadgetItemText(ListIcon_0, selected_row, 0)
  Protected oldValue$ = GetGadgetItemText(ListIcon_0, selected_row, 1)
  Protected newValue$ = ""
  
  If tag$ = "Description"  ;Gebruik EditDescription
    newValue$ = EditDescription(oldValue$)
    newValue$ = ReplaceString(newValue$,Chr(10),"") ;Chr(13) is genoeg!
  ElseIf tag$ = "License"
    newValue$ = EditLicense(oldValue$)
  Else
    newValue$ = InputRequester("Change value", "New value for " + tag$ + ":", oldValue$)
  EndIf
 
  If newValue$ = "" ; gebruiker annuleerde
    ProcedureReturn
  EndIf

  SetGadgetItemText(ListIcon_0, selected_row, newValue$, 1)
EndProcedure



OpenWindow_0()

;======================================== Programma lus ============================================
Global Event, WindowID, GadgetID, Event_Type, Quit.i
Repeat
  Event = WaitWindowEvent() 
  WindowID = EventWindow() 
  GadgetID = EventGadget() 
  Event_Type = EventType() 
  
  Select event
    Case #PB_Event_CloseWindow
      Quit = 1
    Case #PB_Event_Menu
      Select EventMenu()
        Case #MenuItem_2 ;Open
          If Unzip3MF()
            ModelFile$ = UserDir + "Temp_3mf_\3D\3dmodel.model"
            OpenModel(ModelFile$)
          EndIf
        Case #MenuItem_3 ;Save
          If OrcaFile <> ""
            Save3MF(OrcaFile)
          EndIf
        Case #MenuItem_4 ;Save as        
          Save3MF("")
        Case #MenuItem_5 
          Quit = 1
        Case #MenuItem_7
          MessageRequester("About 3MF-Info",infotext)
      EndSelect
    Case #PB_Event_Gadget
      Select EventGadget()
        Case ListIcon_0
          If EventType() = #PB_EventType_LeftDoubleClick
            row = GetGadgetState(ListIcon_0)
            ChangeItem(row)
        EndIf
    EndSelect          
  EndSelect      
Until Quit = 1   

SetCurrentDirectory(startdir) 
Quit = DeleteDirectory(UserDir + "Temp_3mf_\","",#PB_FileSystem_Recursive|#PB_FileSystem_Force) ;alles opruimen
Delay(500)
End


; IDE Options = PureBasic 6.11 LTS (Windows - x64)
; CursorPosition = 1
; Folding = --
; EnableXP
; UseIcon = OrcaFlashForge1.ico
; Executable = ..\3MF-Info.exe
; IncludeVersionInfo
; VersionField0 = 1.0.0.0
; VersionField1 = 1.0.0.0
; VersionField2 = They-Soft
; VersionField3 = 3MF-Info
; VersionField4 = 1.0.0.0
; VersionField5 = 1.0.0.0
; VersionField6 = 3MFmetatag editor
; VersionField7 = 3MF-Info.exe
; VersionField8 = 3MF-Info.exe
; VersionField9 = Freeware
; VersionField17 = 0809 English (United Kingdom)