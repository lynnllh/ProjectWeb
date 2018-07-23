// JavaScript Document

// Example:
// value1 = 3; value2 = 4;
// messageBox("text message %s and %s", value1, value2);
// this message box will display the text "text message 3 and 4"

function messageBox()
{
  var i, msg = "", argNum = 0, startPos;
  var args = messageBox.arguments;
  var numArgs = args.length;
  if(numArgs)
  {
    theStr = args[argNum++];
    startPos = 0;  endPos = theStr.indexOf("%s",startPos);
    if(endPos == -1) endPos = theStr.length;
    while(startPos < theStr.length)
    {
      msg += theStr.substring(startPos,endPos);
      if (argNum < numArgs) msg += args[argNum++];
      startPos = endPos+2;  endPos = theStr.indexOf("%s",startPos);
      if (endPos == -1) endPos = theStr.length;
    }
    if (!msg) msg = args[0];
  }
  alert(msg);
}
