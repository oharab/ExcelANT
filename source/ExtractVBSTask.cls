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

    set writer=nothing
    set classFinder=nothing
  End Sub

End Class