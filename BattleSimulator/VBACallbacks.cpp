#include <assert.h>

#include <Windows.h>

#include <XLCALL.H>

#include "BattleSimulator.h"

#include "VBACallbacks.h"


#if !THREADSAFE
XLOPER12 ActiveWorkbookName(void)
{
    XLOPER12 functionName, result;
    int      returnValue;

    functionName.xltype = xltypeStr;
    functionName.val.str = L"\030ActiveWorkbookNameExport";
    returnValue = Excel12(xlUDF, &result, 1, &functionName);
    assert(returnValue == xlretSuccess);
    assert(result.xltype == xltypeStr);
    return result;
}
#endif


#if !THREADSAFE
XLOPER12 ActiveWorkbookPath(void)
{
    XLOPER12 functionName, result;
    int      returnValue;

    functionName.xltype = xltypeStr;
    functionName.val.str = L"\030ActiveWorkbookPathExport";
    returnValue = Excel12(xlUDF, &result, 1, &functionName);
    assert(returnValue == xlretSuccess);
    assert(result.xltype == xltypeStr);
    return result;
}
#endif


#if !THREADSAFE
void MsgBox(XCHAR promptStr[])
{
    XLOPER12 functionName, prompt, result;
    int      returnValue;

    functionName.xltype = xltypeStr;
    functionName.val.str = L"\014MsgBoxExport";
    prompt.xltype = xltypeStr;
    prompt.val.str = promptStr;
    returnValue = Excel12(xlUDF, &result, 2, &functionName, &prompt);
    assert(returnValue == xlretSuccess);
}
#endif


#if !THREADSAFE
XLOPER12 PathSeparator(void)
{
    XLOPER12 functionName, result;
    int      returnValue;

    functionName.xltype = xltypeStr;
    functionName.val.str = L"\023PathSeparatorExport";
    returnValue = Excel12(xlUDF, &result, 1, &functionName);
    assert(returnValue == xlretSuccess);
    assert(result.xltype == xltypeStr);
    return result;
}
#endif
