class StackItem
  private m_Value

  public property get Value
    Value=m_Value
  End Property
  public property let Value(vValue)
    m_Value=vValue
  End Property

  private m_NextItem

  public property get NextItem
    if isobject(m_NextItem) then
      set NextItem=m_NextItem
    else
      set NextItem=nothing
    end if
  End Property
  public property set NextItem(oNextItem)
    set m_NextItem=oNextItem
  End Property

End Class