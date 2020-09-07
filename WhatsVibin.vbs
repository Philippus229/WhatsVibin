dim objFSO : set objFSO = CreateObject("Scripting.FileSystemObject")
dim objShell : set objShell = CreateObject("WScript.Shell")
dim sessionFile
fileSize = 0
hideSession = false'true
sessionDirectory = "sessions" '\\PCRAUM2\WhatsVibin\sessions
dataDirectory = "data"
absDataDir = objFSO.GetAbsolutePathName(dataDirectory)
currentKey = ""
charAss = ""
current_filepath = ""

function UpdateRelocationTable(protKey)
  protKey = protKey Mod 131071
  if protKey <= 0 then
    protKey = protKey + 131070
  end if
  charAss = ""
  for c0 = 0 to (255 - 32)
    protKey = (protKey Mod 362)^2 'exp = Log(protKey)/Log(10) : (protKey Mod (CInt(2147483647^(1/exp))))^exp
    c1 = 32 + (protKey Mod (256 - 32))
    do while not InStr(charAss, Chr(c1)) = 0
      c1 = 32 + ((c1 - 31) Mod (256 - 32))
    loop
    charAss = charAss & Chr(c1)
  next
end function

function ProcessString(protString, encDec)
  outString = ""
  if encDec = 0 then
    for cIndex = 1 to len(protString)
      outString = outString & Mid(charAss, Asc(Mid(protString, cIndex, 1))-31, 1)
    next
  else
    for cIndex = 1 to len(protString)
      indexc = InStr(charAss, Mid(protString, cIndex, 1))
      if not indexc = 0 then : outString = outString & Chr(indexc + 31)
    next
  end if
  ProcessString = outString
end function

function WriteToSession(session, message)
  if not currentKey = "" then
    message = ProcessString(message, 0)
  end if
  
  objFSO.GetFile(session).Attributes = 0 'show file
  set sessionFile = objFSO.OpenTextFile(session, 8, true, 0)
  sessionFile.WriteLine(message) 'maybe use Write?
  sessionFile.Close
  objFSO.GetFile(session).Attributes = 2*hideSession 'hide file
  set sessionFile = nothing
end function

function ReadFromSession(session, num_msg)
  objFSO.GetFile(session).Attributes = 0 'show file
  set sessionFile = objFSO.OpenTextFile(session, 1)
  sessKey = sessionFile.ReadLine
  content = sessionFile.ReadAll
  sessionFile.Close
  objFSO.GetFile(session).Attributes = 2*hideSession 'hide file
  set sessionFile = nothing
  
  if len(sessKey) > 0 then
    do while (currentKey = "") or (not ProcessString(sessKey, 1) = currentKey)
      currentKey = InputBox("This session is protected! Enter encryption key (integer):", "Protected session!")
      UpdateRelocationTable(CInt(currentKey))
    loop
  else
    currentKey = ""
  end if
  if not currentKey = "" then
    content = ProcessString(content, 1)
  end if
  
  if num_msg > 0 then
    i = 0
    last_index = -1
    do while i < num_msg
      index = inStrRev(content, "<div class='msg'>", last_index)
      if index = 0 then
        exit do
      else
        last_index = index
      end if
      i = i + 1
    loop
    out_string = right(content, len(content)-last_index+1)
    ReadFromSession = out_string
  else
    ReadFromSession = content
  end if
end function

function CreateNewSession(session, setSessActive)
  tmpKey = currentKey
  tmpAss = charAss
  currentKey = InputBox("Choose an encryption key (integer) (leave empty for unprotected session):", "Protect session?")
  'msgbox(currentKey)
  emptyChatFix = " "
  if not currentKey = "" then
    UpdateRelocationTable(CInt(currentKey))
    emptyChatFix = ProcessString(" ", 0)
  end if
  processedKey = ProcessString(currentKey, 0)
  'msgbox(processedKey)
  if not setSessActive then
    currentKey = tmpKey
    charAss = tmpAss
  end if
  
  set sessionFile = objFSO.CreateTextFile(session, true)
  sessionFile.WriteLine processedKey
  sessionFile.Write emptyChatFix
  sessionFile.Close
  set sessionFile = nothing
end function

function UpdateLocalSession(session, num_msg, forced)
  if not objFSO.FileExists(session) then
    CreateNewSession session, true
  end if
  fsize = objFSO.GetFile(session).Size
  if forced = true or not fsize = fileSize then
    fileSize = fsize
    UpdateLocalSession = ReadFromSession(session, num_msg)
  else
    UpdateLocalSession = ""
  end if
end function

'temporary, will probably be replaced by html stuff soon
function Settings(num_msg, autorefresh, keep_inputs)
  num_msg_str = InputBox("Max number of messages to be displayed at once (0=all):", "Number of messages", CStr(num_msg))
  num_msg = CInt(num_msg_str)
  autorefresh_str = InputBox("Delay (milliseconds) between checks (autorefresh if updated) (0=no checks -> no autorefresh):", "New message check delay", CStr(autorefresh))
  autorefresh = CInt(autorefresh_str)
  keep_inputs_str = InputBox("Keep inputs on refresh? (only works for elements with unique ids) (1|0, true|false):", "Keep inputs?", CStr(keep_inputs))
  keep_inputs = CBool(keep_inputs_str)
  Settings = array(num_msg, autorefresh, keep_inputs)
end function

function EmbedFile(filesrc)
  ext = LCase(objFSO.GetExtensionName(filesrc))
  fname = objFSO.GetBaseName(filesrc)
  fileCount = CStr(objFSO.GetFolder(dataDirectory).Files.Count)
  fnout = replace(replace(replace(replace(replace(LCase(Now()),":","")," ",""),".",""),"am",""),"pm","") & fileCount & "." & ext
  filedest = objFSO.BuildPath(dataDirectory, fnout)
  objFSO.CopyFile filesrc, filedest, True
  select case ext
    case "png","jpg","jpeg"
      EmbedFile = "<label class='viewer'><input type='checkbox'>" & _
                  "<div id='preview' filename='" & fnout & "'></div>" & _
                  "<div id ='bg'></div>" & _
                  "<div id='full'></div></label>"
    case else
      EmbedFile = "<a class='unknown-ftype' href='" & fnout & "' download='" & fname & "'>" & _
                  "<div><span id='s0'>&#x1F5CE;</span><span id='s1'>" & fname & " &#x2022; "& UCase(ext) & "</span></div></a>"
  end select
end function

function SelectionHTML(sessionName)
  selectHTML = "<div class='topbar'>Chats</div><div class='chatrooms'>" & _
               "<input type='hidden' id='selected' value='-1'>" & _
               "<button id='chat0' type='submit' onclick='document.all.selected.value=0;'>" & sessionName & "</button>"
  chatIndex = 1
  for each chatFile in objFSO.GetFolder(sessionDirectory).Files
    chatName = Replace(chatFile.Name, ".txt", "")
    if objFSO.GetExtensionName(chatFile.Name) = "txt" and not chatName = sessionName then
      selectHTML = selectHTML + "<button id='chat" & chatIndex & "' type='submit' onclick='document.all.selected.value=" & chatIndex & ";'>" & chatName & "</button>"
      chatIndex = chatIndex + 1
    end if
  next
  selectHTML = selectHTML + "</div>" & _
    "<p class='ctrlbar'><input type='hidden' id='newchat' value='0'>" & _
    "<input class='left' type='submit' value='New' onclick='document.all.newchat.value=1'>" & _
    "<input type='hidden' id='exit' value='0'>" & _
    "<input class='right' type='submit' value='Exit' onclick='document.all.exit.value=1'>" & _
    "<input type='hidden' id='settings' value='0'>" & _
    "<input class='right' type='submit' value='&#x2699;' onclick='document.all.settings.value=1'></p>"
  SelectionHTML = selectHTML
end function

sub MainForm(sessionName, username)
  session = objFSO.BuildPath(sessionDirectory, sessionName & ".txt")
  num_msg = 0 '6
  autorefresh = 1000
  keep_inputs = true
  set ie = CreateObject("InternetExplorer.Application")
  ie.navigate "about:blank"
  while ie.readystate <> 4 : Sleep(100) : wend
  ie.toolbar   = false
  ie.statusbar = false
  iWidth = ie.document.parentWindow.screen.width
  iHeight = ie.document.parentWindow.screen.height
  availDimsX = array(280,480,720,iWidth)
  availDimsY = array(360,800,560,iHeight-28)
  availDimsLeft = array(iWidth-280,CInt(iWidth/2)-240,CInt(iWidth/2)-360,0)
  availDimsTop = array(iHeight-400,CInt(iHeight/2)-400,CInt(iHeight/2)-280,0)
  selected = CInt(InputBox("Choose window size:" & vbNewLine & "0: 280x360" & vbNewLine & "1: 480x800" & vbNewLine & "2: 720x560" & vbNewLine & "3: Maximized", "Window Dimensions", "2"))
  ie.width = availDimsX(selected)
  ie.height = availDimsY(selected)
  ie.left = availDimsLeft(selected)
  ie.top = availDimsTop(selected)
  ie.document.title = "WhatsVibin"
  ie.document.charset = "utf-8" 'i don't even know if this works, but it probably does
  msgs = UpdateLocalSession(session, num_msg, true) 'to disable old commands
  fileSize = 0
  html3 = "<div class='topbar' id='topbar'><input type='hidden' id='back' value='0'><a id='bckbtn' href='#' onclick='document.all.back.value=1'>&#x2b9c;</a>"
  html0 = "</div><div class='msgs' id='msgs'>"
  html1 = "</div>" & _
    "<p class='ctrlbar'><input class='left' type='text' id='msgfield' value='' placeholder='Type a message...'>" & _
    "<input type='hidden' id='s_file' value='0'>" & _
    "<input class='left' type='submit' value='&#x1f4ce;' onclick='document.all.s_file.value=1'>" & _
    "<input type='hidden' id='send' value='0'>" & _
    "<input class='left' type='submit' value='&#x27a4;' onclick='document.all.send.value=1'>" & _
    "<input type='hidden' id='exit' value='0'>" & _
    "<input class='right' type='submit' value='Exit' onclick='document.all.exit.value=1'>" & _
    "<input type='hidden' id='settings' value='0'></p>" 'had to remove settings button because it looked bad
  ie.document.body.innerHTML = html3 & sessionName & html0 & msgs & html1 'update innerHTML
  set style = ie.document.CreateStyleSheet
  style.AddRule "body", "background-color:#075e54;font-family:Helvetica,Helvetica Neue,Arial,sans-serif;font-size:calc(6px + 1.25vh);line-height:calc(10px + 1.25vh);overflow:hidden;"
  style.AddRule ".topbar", "position:relative;width:100%;height:calc(3.125vh + 8px);margin:0;padding-top:calc(1.875vh - 8px);background-color:#075e54;color:white;text-align:center;v-align:center;font-weight:bold;"
  style.AddRule "a#bckbtn", "position:absolute;left:0;top:calc(1.875vh - 8px);color:inherit;text-decoration:inherit;"
  style.AddRule ".msgs", "position:relative;width:(100% - 2px);height:calc(87.5vh - 12px);margin:0;overflow-y:auto;border-radius:4px;border:1px solid white;background-color:#ece5dd;"
  style.AddRule ".msg", "position:relative;width:calc(95% - 10px);height:auto;margin:calc(1.25% + 5px);padding:calc(0.625vh + 4px);background-color:#dcf8c6;border-radius:4px;box-shadow:0 2 4 -1 #6e7c63;"
  style.AddRule ".username", "position:relative;width:100%;height:auto;text-align:left;font-weight:bold;"
  style.AddRule ".message", "position:relative;width:100%;height:auto;" 'width:95%;margin-left:5%;
  style.AddRule ".time", "position:absolute;width:auto;height:auto;font-size:calc(6px + 0.5vh);color:#6e7c63;text-align:right;right:calc(3px + 0.5vh);top:calc(3px + 0.5vh);"
  style.AddRule ".ctrlbar", "position:relative;width:100%;height:calc(10px + 2vh);margin-top:2vh;margin-bottom:2vh;"
  style.AddRule "input#msgfield", "width:calc(80% - 50px);"
  style.AddRule ".left", "float:left;height:100%;font-size:calc(5px + 1.25vh);"
  style.AddRule ".right", "float:right;height:100%;font-size:calc(5px + 1.25vh);"
  style.AddRule "input[type='file']", "display:none;"
  style.AddRule "input", "font-size:inherit;"
  
  style.AddRule "label.viewer > input[type=checkbox]", "display:none;"
  style.AddRule "div#preview", "position:relative;width:33.33vh;height:33.33vh;overflow:hidden;background-size:cover;background-position:center;background-repeat:no-repeat;"
  style.AddRule "div#bg", "display:none;position:fixed;z-index:999;top:calc(5vh + 9px);left:0;width:100%;height:calc(87.5vh - 12px);background-color:rgba(0,0,0,0.5);"
  style.AddRule "label.viewer > input[type=checkbox]:checked ~ div#bg", "display:block;"
  style.AddRule "div#full", "display:none;position:fixed;z-index:1000;top:calc(5vh + 9px);left:0;width:100%;height:calc(87.5vh - 12px);background-position:center;background-repeat:no-repeat;background-size:contain;"
  style.AddRule "label.viewer > input[type=checkbox]:checked ~ div#full", "display:block;"
  
  style.AddRule "a.unknown-ftype", "text-decoration:inherit;width:100%;height:auto;"
  style.AddRule "a.unknown-ftype div", "display:table;width:auto;height:auto;padding:1.25vh;background-color:rgba(0,0,0,0.0625);border:1px solid rgba(0,0,0,0.125);border-radius:2px;"
  style.AddRule "a.unknown-ftype div span#s0", "font-size:calc(18px + 2.5vh);line-height:calc(18px + 2.5vh);margin-right:1.25vh;"
  style.AddRule "a.unknown-ftype div span#s1", "display:table-cell;vertical-align:middle;"
  
  style.AddRule ".chatrooms", "position:relative;width:calc(100% - 2px);height:calc(87.5vh - 12px);margin:0;overflow-y:auto;border-radius:4px;border:1px solid white;background-color:#ece5dd;"
  style.AddRule ".chatrooms button", "position:relative;width:100%;height:auto;padding:1em;border-left:none;border-right:none;border-top:none;border-bottom:1px solid #6e7c63;text-align:left;v-align:center;color:black;background-color:white;text-decoration:none;"
  style.AddRule ".chatrooms button:hover", "background-color:#ece5dd;"
  ie.document.all.msgs.scrollTop = ie.document.all.msgs.scrollHeight-ie.document.all.msgs.clientHeight
  ie.visible = true
  on error resume next
  do while CStr(ie.document.all.exit.value) = "0"
    forceUpdate = false
    if CStr(ie.document.all.back.value) = "1" then
      ie.document.body.innerHTML = SelectionHTML(sessionName)
      do while CStr(ie.document.all.selected.value) = "-1" and CStr(ie.document.all.exit.value) = "0"
        if CStr(ie.document.all.settings.value) = "1" then
          settings_array = Settings(num_msg, autorefresh, keep_inputs)
          num_msg = settings_array(0)
          autorefresh = settings_array(1)
          keep_inputs = settings_array(2)
          ie.document.all.settings.value = 0
        elseif CStr(ie.document.all.newchat.value) = "1" then
          sessName = InputBox("Chatroom name:", "Chatroom name")
          sessPath = objFSO.BuildPath(sessionDirectory, sessName & ".txt")
          if not (sessName = "" or objFSO.FileExists(sessPath)) then
            CreateNewSession sessPath, false
          end if
          ie.document.body.innerHTML = SelectionHTML(sessionName)
        end if
      loop
      if not CStr(ie.document.all.selected.value) = "-1" then
        sessionName = ie.document.all.item("chat" & CStr(ie.document.all.selected.value)).innerText
        session = objFSO.BuildPath(sessionDirectory, sessionName & ".txt")
        forceUpdate = true 'force next update
      elseif CStr(ie.document.all.exit.value) = "1" then
        exit do
      end if
    end if
    msgs = UpdateLocalSession(session, num_msg, forceUpdate) 'check for update
    if len(msgs) > 0 then 'refresh if updated
      set valueDict = CreateObject("Scripting.Dictionary")
      if keep_inputs then 'very temporary, next version probably won't require (unique) ids
        old_html = ie.document.body.innerHTML
        input_index = InStr(old_html, "<input ")
        do while input_index > 0
          input_id = Mid(old_html, InStr(input_index, old_html, "id=")+4)
          input_id = Replace(input_id, """", "'")
          input_id = Left(input_id, InStr(input_id, "'")-1)
          valueDict.Add input_id, ie.document.all.item(input_id).value
          input_index = InStr(input_index+1, old_html, "<input ")
        loop
      end if
      old_count = 0
      for each elem in ie.document.getElementsByClassName("user-script")
        old_count = old_count + 1
      next
      scrollTop = ie.document.all.msgs.scrollTop
      scrollBottom = ie.document.all.msgs.scrollHeight-ie.document.all.msgs.scrollTop-ie.document.all.msgs.clientHeight
      ie.document.body.innerHTML = html3 & sessionName & html0 & msgs & html1 'update innerHTML
      for each viewer in ie.document.getElementsByClassName("viewer")
        new_path = "url('" & replace(objFSO.BuildPath(absDataDir, viewer.children.item("preview").getAttribute("filename")),"\","/") & "')"
        'if objFSO.FileExists(new_path) then
        viewer.children.item("preview").style.backgroundImage = new_path
        viewer.children.item("full").style.backgroundImage = new_path
        'end if
      next
      for each unknown_ftype in ie.document.getElementsByClassName("unknown-ftype")
        new_path = objFSO.BuildPath(absDataDir, unknown_ftype.getAttribute("href"))
        unknown_ftype.setAttribute "href", new_path
      next
      if scrollBottom = 0 then
        ie.document.all.msgs.scrollTop = ie.document.all.msgs.scrollHeight-ie.document.all.msgs.clientHeight
      else
        ie.document.all.msgs.scrollTop = scrollTop
      end if
      for each in_id in valueDict
        if not ie.document.all.item(in_id) = nothing then
          ie.document.all.item(in_id).value = valueDict(in_id)
        end if
      next
      set valueDict = nothing
      e_index = 0
      for each elem in ie.document.getElementsByClassName("user-script") 'this is terrible but i don't know a better way to do it without unique ids
        elem.children.item("run").value = 0
        e_index = e_index + 1
        if e_index = old_count then
          exit for
        end if
      next
    end if
    break = 0
    do while break = 0
      break = autorefresh+ie.document.all.send.value+ie.document.all.exit.value+ie.document.all.settings.value+ie.document.all.back.value
      if CStr(ie.document.all.send.value) = "1" then 'send message/file
        msg = ie.document.all.msgfield.value
        if current_filepath <> "" then
          msg = EmbedFile(current_filepath) & msg
          current_filepath = ""
        end if
        if not msg = "" then
          if InStr(msg, "!") = 1 then
            cmd = LCase(Mid(msg, 2, InStr(msg, " ")-2))
            args = trim(right(msg, len(msg)-InStr(msg, " ")))
            select case cmd
              case "button", "field", "run"
                lang = left(args, InStr(args, " ")-1)
                cont = right(args, len(args)-InStr(args, " "))
                if cmd <> "run" then : msg = "" : end if
                msg = msg & "<div class='user-script'><input type='hidden' id='lang' value='" & lang & "'><input type='"
                if cmd = "field" then : msg = msg & "text" : else : msg = msg & "hidden" : end if
                msg = msg & "' id='content' value='" & cont & "' placeholder='Enter code/command...'><input type='hidden' id='run' value='"
                if cmd = "run" then
                  msg = msg & "1'></div>"
                else
                  if cmd = "field" then : label = ">>" : else : label = cont : end if
                  msg = msg & "0'><input type='submit' value='" & label & "' onclick='this.parentNode.children.item(""run"").value = 1;'></div>"
                end if
            end select
          end if
          WriteToSession session, "<div class='msg'><div class='username'>" & username & _
                                  "</div><div class='time'>" & Now() & _
                                  "</div><div class='message'>" & msg & "</div></div>"
        end if
        ie.document.all.send.value = 0
        ie.document.all.msgfield.value = ""
      elseif CStr(ie.document.all.s_file.value) = "1" then
        set objExec = objShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
        current_filepath = objExec.stdout.ReadLine()
        ie.document.all.s_file.value = 0
      elseif CStr(ie.document.all.settings.value) = "1" then 'open settings
        settings_array = Settings(num_msg, autorefresh, keep_inputs)
        num_msg = settings_array(0)
        autorefresh = settings_array(1)
        keep_inputs = settings_array(2)
        ie.document.all.settings.value = 0
      end if
      for each elem in ie.document.getElementsByClassName("user-script") 'check for/run user-script
        lang = LCase(elem.children.item("lang").value)
        cont = elem.children.item("content").value
        if CStr(elem.children.item("run").value) = "1" then
          select case lang
            case "vbscript"
              Execute(cont)
            case "shell"
              objShell.Run(cont)
          end select
          elem.children.item("run").value = 0
        end if
      next
    loop
    Sleep(autorefresh)
  loop
  ie.Quit
end sub

if not objFSO.FolderExists(sessionDirectory) then
  objFSO.CreateFolder(sessionDirectory)
end if

if not objFSO.FolderExists(dataDirectory) then
  objFSO.CreateFolder(dataDirectory)
end if

defaultUsername = CreateObject("WScript.Network").UserName
MainForm InputBox("Chatroom name:", "Create/join chatroom", "Info 10"), InputBox("Enter username:", "Username", defaultUsername)
