#pragma once


#include <XLCALL.H>


/* implemented as a macro so a variable number of XLOPER12 structures can be freed in one xlFree call */
#define FREE(n, ...)                                        \
{                                                           \
    int returnValue;                                        \
                                                            \
    returnValue = Excel12(xlFree, nullptr, n, __VA_ARGS__); \
    assert(returnValue == xlretSuccess);                    \
}


bool     GetNamedBoolean (XCHAR nameStr[]);

double   GetNamedNumber  (XCHAR nameStr[]);

XLOPER12 GetNamedRange   (XCHAR nameStr[]);

XLOPER12 GetNamedArray   (XCHAR nameStr[]);


double   IndexNumber     (const XLOPER12 &arrayRange, int rowNum, int colNum);


int      Match           (double lookupValue, const XLOPER12 &lookupRange, int matchType);

int      Match           (const XLOPER12 &lookupStr, const XLOPER12 &lookupRange, int matchType);


double   VLookupNumber   (double lookupValue, const XLOPER12 &lookupRange, int colNum, bool approximateMatch);

XLOPER12 VLookupString   (double lookupValue, const XLOPER12 &lookupRange, int colNum, bool approximateMatch);


#if 0
/* for reference */
void     Free            (XLOPER12 &operand);
#endif
