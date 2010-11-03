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
    wscript.echo "Storing " & sFileName
    m_files.Push sFileName
  End Sub

  Public Sub Execute
    if m_fso.FileExists(me.File) then
      m_fso.DeleteFile(me.File)
    end if
    dim op
    dim f
    f=m_files.Pop

    set op=m_fso.OpenTextFile(me.File,8,true)
    do until f=""
      dim sFile
      Wscript.echo "Reading " & f
      sFile=m_fso.OpenTextFile(f).ReadAll
      op.Write sFile
      f=m_files.Pop
    loop 

    op.Close
    set op=Nothing
  End Sub

End Class