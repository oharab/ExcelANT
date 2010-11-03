class BuildVBSTask
  private m_fso
  private m_files

  Private Sub Class_Initialize()
    set m_fso=CreateObject("Scripting.FileSystemObject")
    set m_files=new Stack
  End Sub

  Private Sub Class_Terminate()
    set m_fso=nothing
    set m_files=nothing
  End Sub

  private m_File

  public property get File
    File=m_File
  End Property
  public property let File(sFile)
    m_File=sFile
  End Property

  Public Sub AddInput(sFileName)
    m_files.Push sFileName
  End Sub

  Public Sub Execute
    dim op
    set op=m_fso.OpenTextFile(me.File,2)
    dim f
    f=m_files.Pop

    do until f=""
      Wscript.echo f
      f=m_files.Pop
      op.Write f
    loop 

    op.Close
    set op=Nothing
  End Sub

End Class