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