class Stack
  private m_current

  Public Sub Push(vValue)
    dim si
    set si=new StackItem
    si.Value=vValue
    if isobject(m_current) then
      set si.NextItem=m_current
    end if
    set m_current=si
  End Sub

  Public Function Pop
    if not m_current is nothing then
      Pop=m_current.Value
      set m_current=m_current.NextItem
    end if
  End Function

  Public Function Peek
    if not m_current is nothing then
      Peek=m_current.Value
    end if
  End Function

End Class