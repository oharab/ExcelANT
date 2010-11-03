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

    if m_opts.Exists("target") then
      m_target=m_opts.Value("target")
    end if
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
    " -buildfile=<text>" & vbtab & "Use given buildfile " & vbcrlf & _
    " -target=<text>" & vbtab & "Process given target" & vbcrlf & _
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