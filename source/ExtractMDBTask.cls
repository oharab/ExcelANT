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