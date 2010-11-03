class Options
  private m_opts

  Private Sub Class_Initialize()
    set m_opts=CreateObject("Scripting.Dictionary")

    dim argumentCount, splitter,remover
    set splitter=new RegExp
    with splitter
      .IgnoreCase=true
      .Global=true
      .Pattern="^-{1,2}|^/|="
    end with
    set remover=new RegExp
    with remover
      .IgnoreCase=true
      .Global=true
      .Pattern="^['""]?(.*?)['""]?$"
    end with
    dim txt,argument,matches,parts,part,opt
    for each txt in WScript.Arguments
      parts=split(splitter.Replace(txt,"####è"),"####è",3)
      select case ubound(parts)
        case 0
          if argument <>"" then
            if not m_opts.Exists(argument) then
              Parts(0) =Remover.Replace(Parts(0), "$1")
              m_opts.Add argument, parts(0)
            end if
            argument = ""
          end if
          ' else Error: no parameter waiting for a value (skipped)
        case 1
          ' The last parameter is still waiting. 
          ' With no value, set it to true.

          if argument<>"" then
            if not m_opts.Exists(argument) then
              m_opts.Add argument, true
            end if
          end if
          argument = Parts(1)

          'Parameter with enclosed value
        case 2
          ' The last parameter is still waiting. 
          ' With no value, set it to true.
          if argument<>"" then
            if not m_opts.Exists(argument) then 
              m_opts.Add argument,true
            end if
          end if

          argument=parts(1)
          ' Remove possible enclosing characters (",')
          if not m_opts.Exists(argument) then
            parts(2) = remover.Replace(Parts(2), "$1")
            m_opts.Add argument, parts(2)
          end if
          argument=""                        
      end select
    next

    if argument<>"" then
      if not m_opts.Exists(argument) then
        m_opts.Add argument,true
      end if
    end if
  end sub
  Private Sub Class_Terminate()
    set m_opts=nothing
  end sub

  public Property Get Count
    Count=m_opts.Count 
  end property

  public function Exists(sName)
    Exists=m_opts.Exists(sName)
  end function

  public Function Value(sName)
    if m_opts.Exists(sName) then
      Value=m_opts.Item(sName)
    else
      Value=false
    end if
  end function

  public property get Keys
    Keys=m_opts.Keys
  end property
end class