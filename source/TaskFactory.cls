class TaskFactory
  private m_BuildPath
  private m_fso

  private sub Class_Initialize()
    set m_fso=CreateObject("Scripting.FileSystemObject")
  end sub
  private sub Class_Terminate()
    set m_fso=nothing
  end sub

  public property get BuildPath
    BuildPath=m_BuildPath
  End Property
  public property let BuildPath(sBuildPath)
    m_BuildPath=sBuildPath
  End Property

  public function Build(taskNode)
    select case taskNode.nodeName
      case "print"
        set Build=new PrintTask
        Build.Text=taskNode.attributes.getNamedItem("text").nodeValue
      case "printValues"
        set Build=new PrintValuesTask
        dim value
        for each value in taskNode.getElementsByTagName("text")
          Build.Add value.attributes.getNamedItem("value").nodeValue
        next
      case "extractmdb"
        set Build=new ExtractMDBTask
        Build.Database=m_fso.GetFile(m_fso.BuildPath(Me.BuildPath,taskNode.attributes.getNamedItem("database").nodeValue)).Path
        Build.Output=m_fso.GetFolder(m_fso.BuildPath(Me.BuildPath,taskNode.attributes.getNamedItem("output").nodeValue)).Path
        dim wrkgNode
        for each wrkgNode in taskNode.getElementsByTagName("wrkgrp")
          Build.Workgroup=m_fso.GetFile(m_fso.BuildPath(Me.BuildPath,wrkgNode.getElementsByTagName("mdw")(0).text)).Path
          Build.Username=wrkgNode.getElementsByTagName("username")(0).text
          Build.Password=wrkgNode.getElementsByTagName("password")(0).text
        next
      case "extractVBS"
        set Build=new ExtractVBSTask
        Build.File=m_fso.GetFile(m_fso.BuildPath(Me.BuildPath,taskNode.attributes.getNamedItem("file").nodeValue)).Path
        Build.Output=m_fso.GetFolder(m_fso.BuildPath(Me.BuildPath,taskNode.attributes.getNamedItem("output").nodeValue)).Path
      case "buildVBS"
        set Build=new BuildVBSTask
        Build.File=m_fso.BuildPath(Me.BuildPath,taskNode.attributes.getNamedItem("file").nodeValue)
        dim chd
        for each chd in taskNode.childNodes
          select case chd.tagName
            case "file"
              Build.AddInput m_fso.BuildPath(Me.BuildPath,chd.childNodes(0).nodeValue)
            case "filecollection"
              dim path,parent,filter
              path=m_fso.BuildPath(Me.BuildPath,chd.childNodes(0).nodeValue)
              if m_fso.FolderExists(path) then
                parent=path
                filter = "*.*"
              else
                parent=m_fso.GetParentFolderName(path)
                if parent="" then 
                  if right(path,1)=":" then
                    parent=path
                  else
                    parent="."
                  end if
                end if
                filter=m_fso.GetFileName(path)
                if filter="" then Filter="*.*"
              end if
              dim file
              for each file in m_fso.GetFolder(parent).Files
                if CompareFileName(File.Name,filter) then
                  Build.AddInput file.Path
                end if
              next
              
            case else
              err.raise -1001,"Unknown  option '" & chd.tagName & "' in buildVBS."
          end select
        next
      case else
        err.raise 1,"Unknown Task " & taskNode.nodeName
    end select
  end function

  private function CompareFileName(byval name,byval filter)
    Common=false
    dim np,fp
    np=1
    fp=1
    do
      if fp>len(filter) then
        CompareFileName=np>len(name)
        exit function
      end if
If Mid(Filter,fp) = ".*" Then    ' special case: ".*" at end of filter
         If np > Len(Name) Then CompareFileName = True: Exit Function
         End If
      If Mid(Filter,fp) = "." Then     ' special case: "." at end of filter
         CompareFileName = np > Len(Name): Exit Function
         End If
      Dim fc: fc = Mid(Filter,fp,1): fp = fp + 1
      Select Case fc
         Case "*"
            CompareFileName = CompareFileName2(name,np,filter,fp)
            Exit Function
         Case "?"
            If np <= Len(Name) And Mid(Name,np,1) <> "." Then np = np + 1
         Case Else
            If np > Len(Name) Then Exit Function
            Dim nc: nc = Mid(Name,np,1): np = np + 1
            If Strcomp(fc,nc,vbTextCompare)<>0 Then Exit Function
         End Select
      Loop
   End Function

Private Function CompareFileName2 (ByVal Name, ByVal np0, ByVal Filter, ByVal fp0)
   Dim fp: fp = fp0
   Dim fc2
   Do                                  ' skip over "*" and "?" characters in filter
      If fp > Len(Filter) Then CompareFileName2 = True: Exit Function
      fc2 = Mid(Filter,fp,1): fp = fp + 1
      If fc2 <> "*" And fc2 <> "?" Then Exit Do
      Loop
   If fc2 = "." Then
      If Mid(Filter,fp) = "*" Then     ' special case: ".*" at end of filter
         CompareFileName2 = True: Exit Function
         End If
      If fp > Len(Filter) Then         ' special case: "." at end of filter
         CompareFileName2 = InStr(np0,Name,".") = 0: Exit Function
         End If
      End If
   Dim np
   For np = np0 To Len(Name)
      Dim nc: nc = Mid(Name,np,1)
      If StrComp(fc2,nc,vbTextCompare)=0 Then
         If CompareFileName(Mid(Name,np+1),Mid(Filter,fp)) Then
            CompareFileName2 = True: Exit Function
            End If
         End If
      Next
   CompareFileName2 = False
   End Function
        
end class