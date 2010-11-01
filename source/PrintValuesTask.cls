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