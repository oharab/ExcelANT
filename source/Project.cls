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

  public sub Execute()
    m_targets.item(Me.Default).Execute
  end sub
end class