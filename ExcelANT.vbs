Option Explicit
const version="0.0.1"

dim ea
set ea=new ExcelANT
ea.CreateProject
ea.Execute
set ea=nothing

class ExcelANT
  private m_fso
  private m_opts
  private m_buildfile
  private m_currentdir
  private m_project
  private m_target

  Private Sub Class_Initialize()
    set m_fso=CreateObject("Scripting.FileSystemObject")
    set m_opts=new Options

    'Check for help parameter
    if m_opts.Exists("help") or m_opts.Exists("h") then
      me.showUsage()
    end if

    m_currentdir=m_fso.GetAbsolutePathName(".")

    if m_opts.Exists("buildfile") then
      m_buildfile=m_opts.Value("buildfile")
    end if

    WScript.echo "Last Argument:" & WScript.Arguments(Wscript.Arguments.count-1) 
    WScript.echo "Build exists:" & m_opts.Exists("build")
  end Sub 

  Private Sub Class_Terminate()
    set m_project=nothing
    set m_buildfile=nothing
    set m_opts=nothing
    set m_fso=nothing
  end sub

  private function showUsage()
    WScript.Echo "ExcelANT " & version & vbcrlf & _
    "Copyright (C) 2010 bpo@robotparade.co.uk" & vbcrlf & _
    vbcrlf & _
    "Usage:  " & _
    vbtab & WScript.ScriptName & " [options] <target> <target>" & vbcrlf & _
    "Options:" & vbcrlf & _
    vbcrlf & _
    " -buildfile:<text>" & vbtab & "Use given buildfile (Short format /f)" & vbcrlf & _
    " -h[elp]" & vbtab & "Prints this message" & vbcrlf & _
    vbcrlf & _
    "A file ending in .build will be used if no buildfile is specified."
    Wscript.Quit
  end function

  public property get BuildFile()
    if m_buildfile = "" then
      if m_fso.FileExists(m_fso.BuildPath(m_currentdir,"default.build")) then
        m_buildfile=m_fso.BuildPath(m_currentdir,"default.build")
      else
        dim f
        for each f in m_fso.GetFolder(m_currentdir).Files
          if right(f.Name,6)=".build" then
            m_buildfile=f.Path
            exit for
          end if
        next
      end if
    else
      if m_fso.FileExists(m_buildfile) then
      else
        if m_fso.FileExists(m_fso.BuildPath(m_currentdir,m_buildfile)) then
          m_buildfile=m_fso.BuildPath(m_currentdir,m_buildfile)
        else
          m_buildfile=""
        end if
      end if
    end if
    
    if m_buildfile="" then
      showUsage()
    else
      BuildFile=m_buildfile
    end if
  end property

  public Sub CreateProject()
    dim doc
    set doc=CreateObject("Microsoft.xmlDOM")
    if not doc.Load(me.buildfile) then showInvalidBuildFile(doc)
    dim projectNode
    set projectNode=doc.childNodes(1)
    if projectNode.NodeName<>"project" then showUsage
    set m_project=new Project
    with m_project
      .Name=projectNode.attributes.getNamedItem("name").nodeValue
      .Default=projectNode.attributes.getNamedItem("default").nodeValue
      .BuildPath=m_fso.GetFile(Me.BuildFile).ParentFolder.Path
    end with

    dim projectPropertyNode
    for each projectPropertyNode in projectNode.getElementsByTagName("property")
      with projectPropertyNode.attributes
        m_project.AddProperty .getNamedItem("name").nodeValue,.getNamedItem("value").nodeValue
      end with
    next
    dim targetNode
    for each targetNode in projectNode.getElementsByTagName("target")
      m_project.AddTarget targetNode
    next
    
    set doc=nothing
  end sub

  private sub showInvalidBuildFile(doc)
    with doc.parseError
      Wscript.echo "Your Build file failed to load" & _
      "due the following error." & vbCrLf & _
      "Error #: " & .errorCode & ": " & .reason & _
      "Line #: " & .Line & vbCrLf & _
      "Line Position: " & .linepos & vbCrLf & _
      "Position In File: " & .filepos & vbCrLf & _
      "Source Text: " & .srcText
    end with
    showUsage
  end sub

  public Sub Execute()
    m_project.Execute m_target
  end sub
end class

class Target
  private m_name
  private m_tasks
  private m_taskFactory

  private Sub Class_Initialize
    set m_tasks=CreateObject("Scripting.Dictionary")
    set m_taskFactory=new TaskFactory
  end sub

  private sub Class_Terminate
    set m_taskFactory=Nothing
    set m_tasks=nothing
  end sub

  public Sub AddTasks(taskNodes)
    dim nd
    for each nd in taskNodes
      m_tasks.Add nd.NodeName,m_taskFactory.Build(nd)
    next
  end sub

  public Property get Name
    Name=m_name
  end property
  public property let Name(sName)
    m_name=sname
  end property

  public Sub Execute()
    Wscript.Echo "Executing " & me.Name
    dim task
    for each task in m_tasks.items
      task.Execute
    next
  end sub

  private m_BuildPath

  public property get BuildPath
    BuildPath=m_BuildPath
  End Property
  public property let BuildPath(sBuildPath)
    m_BuildPath=sBuildPath
    m_taskFactory.BuildPath=sBuildPath
  End Property

end class

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

class ExtractVBSTask
  private m_fso

  Private Sub Class_Initialize()
    set m_fso=CreateObject("Scripting.FileSystemObject")
  End Sub

  Private Sub Class_Terminate()
    set m_fso=nothing
  End Sub

  private m_File

  public property get File
    File=m_File
  End Property
  public property let File(sFile)
    m_File=sFile
  End Property

  private m_Output

  public property get Output
    Output=m_Output 
  End Property
  public property let Output(sOutput)
    m_Output=sOutput
  End Property

  
  Public Sub Execute
    dim classFinder
    WScript.echo "Saving to " & me.Output

    set classFinder=new RegExp
    with classFinder
      .IgnoreCase = True
      .Global=True
      .MultiLine=True
      .pattern="^class (.+?)$[\s\S]+?end class$"
    end with

    dim sFile
    sFile=m_fso.OpenTextFile(Me.File).ReadAll

    dim m,writer,filename
    for each m in classFinder.Execute(sFile)
      fileName=m_fso.BuildPath(Me.Output,m.SubMatches.item(0)) & ".cls"
      set writer=m_fso.OpenTextFile(fileName,2,true)
      writer.Write m.Value
    next

    dim mainText
    mainText=classFinder.replace(sFile,"")

    fileName=m_fso.BuildPath(Me.Output,"main.bas")
    set writer=m_fso.OpenTextFile(fileName,2,true)
    writer.Write mainText

    set writer=nothing
    set classFinder=nothing
  End Sub

End Class


class ExtractMDBTask
  private m_fso

  private sub Class_Initialize()
    set m_fso=CreateObject("Scripting.FileSystemObject")
  end sub
  private sub Class_Terminate()
    set m_fso=nothing
  end sub

  private m_Database
  Public Property Get Database
    Database=m_Database
  End Property
  Public Property Let Database(sDatabase)
    m_Database=sDatabase
  End Property

  private m_Workgroup

  public property get Workgroup
    Workgroup=m_Workgroup
  End Property
  public property let Workgroup(sWorkgroup)
    m_Workgroup=sWorkgroup
  End Property

  private m_UserName

  public property get UserName
    UserName=m_UserName
  End Property
  public property let UserName(sUserName)
    m_UserName=sUserName
  End Property

  private m_Password

  public property get Password
    Password=m_Password
  End Property
  public property let Password(sPassword)
    m_Password=sPassword
  End Property

  private m_Output

  public property get Output
    Output=m_Output
  End Property
  public property let Output(sOutput)
    m_Output=sOutput
  End Property

  
  public sub Execute()
    dim app,dbengine,ws,db
    'May have to do a check for different versions here
    set app=CreateObject("Access.Application")
    wscript.echo app.DBEngine.SystemDB
    app.OpenCurrentDatabase "O:\Common\dev\SurveyProcessing\build\Me.Database",false
    dim m
    for each m in app.CurrentProject.AllModules
      Wscript.echo m.Name
    next
    set app=nothing
  end sub

end class

class PrintTask
  private m_text

  public sub Execute()
    Wscript.Echo m_text
  end sub

  public property Let Text(sText)
    m_text=sText
  end property

end class

class PrintValuesTask
  private m_values

  private sub Class_Initialize()
    set m_values=CreateObject("scripting.Dictionary")
  end sub
  
  private sub Class_Terminate()
    set m_values=nothing
  end sub

  public sub Execute()
    dim value
    for each value in m_values.keys
      Wscript.echo value
    next
  end sub

  public Sub Add(sValue)
    m_values.add sValue,sValue
  end sub
  
end class

class Project
  private m_properties
  private m_name
  private m_default
  private m_targets

  private Sub Class_Initialize()
    set m_properties=CreateObject("Scripting.Dictionary")
    set m_targets=CreateObject("Scripting.Dictionary")
  end sub

  private sub Class_Terminate()
    set m_properties=nothing
    set m_targets=nothing
  end sub

  private m_BuildPath

  public property get BuildPath
    BuildPath=m_BuildPath
  End Property
  public property let BuildPath(sBuildPath)
    m_BuildPath=sBuildPath
  End Property

  public property get Name
    Name=m_name
  end property
  public property let Name(sName)
    m_name=sName
  end property

  public property get Default
    Default=m_default
  end property
  public property let Default(sDefault)
    m_default=sDefault
  end property

  public Sub AddProperty(sName,sValue)
    m_properties.Add sName,sValue
  end sub

  public sub AddTarget(targetNode)
    dim t,targetName
    targetName=targetNode.attributes.getNamedItem("name").nodeValue
    m_targets.Add targetName,new Target
    set t=m_targets.item(targetName)
    t.Name=targetName 
    t.BuildPath=me.BuildPath
    t.AddTasks targetNode.childNodes
  end sub

  public sub Execute(targetName)
    if targetName="" then targetName=Me.Default
    WScript.Echo targetName
    m_targets.item(targetName).Execute
  end sub
end class

class Options
  private m_opts

  Private Sub Class_Initialize()
    set m_opts=CreateObject("Scripting.Dictionary")

    dim argumentCount, splitter,remover
    set splitter=new RegExp
    with splitter
      .IgnoreCase=true
      .Global=true
      .Pattern="^-{1,2}|^/|="
    end with
    set remover=new RegExp
    with remover
      .IgnoreCase=true
      .Global=true
      .Pattern="^['""]?(.*?)['""]?$"
    end with
    dim txt,argument,matches,parts,part,opt
    for each txt in WScript.Arguments
      wscript.echo "Working on argument:" & txt
      parts=split(splitter.Replace(txt,"####è"),"####è",3)
      select case ubound(parts)
        case 0
          if argument <>"" then
            if not m_opts.Exists(argument) then
              Parts(0) =Remover.Replace(Parts(0), "$1")
              m_opts.Add argument, parts(0)
            end if
            argument = ""
          end if
          ' else Error: no parameter waiting for a value (skipped)
        case 1
          ' The last parameter is still waiting. 
          ' With no value, set it to true.

          if argument<>"" then
            if not m_opts.Exists(argument) then
              m_opts.Add argument, true
            end if
          end if
          argument = Parts(1)

          'Parameter with enclosed value
        case 2
          ' The last parameter is still waiting. 
          ' With no value, set it to true.
          if argument<>"" then
            if not m_opts.Exists(argument) then 
              m_opts.Add argument,true
            end if
          end if

          argument=parts(1)
          ' Remove possible enclosing characters (",')
          if not m_opts.Exists(argument) then
            parts(2) = remover.Replace(Parts(2), "$1")
            m_opts.Add argument, parts(2)
          end if
          argument=""                        
      end select
    next

    if argument<>"" then
      if not m_opts.Exists(argument) then
        m_opts.Add argument,true
      end if
    end if
  end sub
  Private Sub Class_Terminate()
    set m_opts=nothing
  end sub

  public Property Get Count
    Count=m_opts.Count 
  end property

  public function Exists(sName)
    Exists=m_opts.Exists(sName)
  end function

  public Function Value(sName)
    if m_opts.Exists(sName) then
      Value=m_opts.Item(sName)
    else
      Value=false
    end if
  end function

  public property get Keys
    Keys=m_opts.Keys
  end property
end class
