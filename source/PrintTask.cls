class PrintTask
  private m_text

  public sub Execute()
    Wscript.Echo m_text
  end sub

  public property Let Text(sText)
    m_text=sText
  end property

end class