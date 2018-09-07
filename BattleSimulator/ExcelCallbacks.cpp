#include <assert.h>

#include <Windows.h>

#include <XLCALL.H>

#include "ExcelCallbacks.h"


bool GetNamedBoolean(XCHAR nameStr[])
{
    XLOPER12 name, evaluateResult, coerceResult;
    int      returnValue;

    name.xltype = xltypeStr;
    name.val.str = nameStr;
    returnValue = Excel12(xlfEvaluate, &evaluateResult, 1, &name);
    assert(returnValue == xlretSuccess);
    assert(evaluateResult.xltype == xltypeRef);
    /* look up cell value from reference */
    returnValue = Excel12(xlCoerce, &coerceResult, 1, &evaluateResult);
    FREE(1, &evaluateResult);
    assert(returnValue == xlretSuccess);
    assert(coerceResult.xltype == xltypeBool);
    return coerceResult.val.xbool != 0;
}


double GetNamedNumber(XCHAR nameStr[])
{
    XLOPER12 name, evaluateResult, coerceResult;
    int      returnValue;

    name.xltype = xltypeStr;
    name.val.str = nameStr;
    returnValue = Excel12(xlfEvaluate, &evaluateResult, 1, &name);
    assert(returnValue == xlretSuccess);
    assert(evaluateResult.xltype == xltypeRef);
    /* look up cell value from reference */
    returnValue = Excel12(xlCoerce, &coerceResult, 1, &evaluateResult);
    FREE(1, &evaluateResult);
    assert(returnValue == xlretSuccess);
    assert(coerceResult.xltype == xltypeNum);
    return coerceResult.val.num;
}


XLOPER12 GetNamedRange(XCHAR nameStr[])
{
    XLOPER12 name, result;
    int      returnValue;

    name.xltype = xltypeStr;
    name.val.str = nameStr;
    returnValue = Excel12(xlfEvaluate, &result, 1, &name);
    assert(returnValue == xlretSuccess);
    assert(result.xltype == xltypeRef);
    return result;
}


XLOPER12 GetNamedArray(XCHAR nameStr[])
{
    XLOPER12 name, evaluateResult, coerceResult;
    int      returnValue;

    name.xltype = xltypeStr;
    name.val.str = nameStr;
    returnValue = Excel12(xlfEvaluate, &evaluateResult, 1, &name);
    assert(returnValue == xlretSuccess);
    assert(evaluateResult.xltype == xltypeRef);
    /* look up cell values from reference */
    returnValue = Excel12(xlCoerce, &coerceResult, 1, &evaluateResult);
    FREE(1, &evaluateResult);
    assert(returnValue == xlretSuccess);
    assert(coerceResult.xltype == xltypeMulti);
    return coerceResult;
}


double IndexNumber(const XLOPER12 &arrayRange, int rowNum, int colNum)
{
    XLOPER12 row, column, indexResult, coerceResult;
    int      returnValue;

    row.xltype = xltypeInt;
    row.val.w = rowNum;
    column.xltype = xltypeInt;
    column.val.w = colNum;
    returnValue = Excel12(xlfIndex, &indexResult, 3, &arrayRange, &row, &column);
    assert(returnValue == xlretSuccess);
    assert(indexResult.xltype == xltypeRef);
    /* look up cell value from reference */
    returnValue = Excel12(xlCoerce, &coerceResult, 1, &indexResult);
    FREE(1, &indexResult);
    assert(returnValue == xlretSuccess);
    assert(coerceResult.xltype == xltypeNum || coerceResult.xltype == xltypeNil);
    if (coerceResult.xltype == xltypeNil) {
        return 0.0;
    } else {
        return coerceResult.val.num;
    }
}


int Match(double lookupValue, const XLOPER12 &lookupRange, int matchType)
{
    XLOPER12 value, match, result;
    int      returnValue;

    value.xltype = xltypeNum;
    value.val.num = lookupValue;
    match.xltype = xltypeInt;
    match.val.w = matchType;
    returnValue = Excel12(xlfMatch, &result, 3, &value, &lookupRange, &match);
    assert(returnValue == xlretSuccess);
    assert(result.xltype == xltypeNum);
    return (int) result.val.num;
}


int Match(const XLOPER12 &lookupStr, const XLOPER12 &lookupRange, int matchType)
{
    XLOPER12 match, result;
    int      returnValue;

    match.xltype = xltypeInt;
    match.val.w = matchType;
    returnValue = Excel12(xlfMatch, &result, 3, &lookupStr, &lookupRange, &match);
    assert(returnValue == xlretSuccess);
    assert(result.xltype == xltypeNum);
    return (int) result.val.num;
}


double VLookupNumber(double lookupValue, const XLOPER12 &lookupRange, int colNum, bool approximateMatch)
{
    XLOPER12 value, column, approximate, result;
    int      returnValue;

    value.xltype = xltypeNum;
    value.val.num = lookupValue;
    column.xltype = xltypeInt;
    column.val.w = colNum;
    approximate.xltype = xltypeBool;
    approximate.val.xbool = approximateMatch;
    returnValue = Excel12(xlfVlookup, &result, 4, &value, &lookupRange, &column, &approximate);
    assert(returnValue == xlretSuccess);
    assert(result.xltype == xltypeNum);
    return result.val.num;
}


XLOPER12 VLookupString(double lookupValue, const XLOPER12 &lookupRange, int colNum, bool approximateMatch)
{
    XLOPER12 value, column, approximate, result;
    int      returnValue;

    value.xltype = xltypeNum;
    value.val.num = lookupValue;
    column.xltype = xltypeInt;
    column.val.w = colNum;
    approximate.xltype = xltypeBool;
    approximate.val.xbool = approximateMatch;
    returnValue = Excel12(xlfVlookup, &result, 4, &value, &lookupRange, &column, &approximate);
    assert(returnValue == xlretSuccess);
    assert(result.xltype == xltypeStr || result.xltype == xltypeNil);
    return result;
}


/* every xltypeStr, xltypeRef, and xltypeMulti returned from Excel12 must be freed */
#if 0
/* for reference */
void Free(XLOPER12 &operand)
{
    int returnValue;

    returnValue = Excel12(xlFree, nullptr, 1, &operand);
    assert(returnValue == xlretSuccess);
}
#endif
