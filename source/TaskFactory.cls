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
      case else
        err.raise 1,"Unknown Task " & taskNode.nodeName
    end select
  end function
end class