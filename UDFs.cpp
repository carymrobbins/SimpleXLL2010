#include <windows.h>
#include <xlcall.h>
#include <framewrk.h>
#include <boost/preprocessor.hpp>

#define VARG_MAX 200
#define VARG_FORMAT(Z, A, B) BOOST_PP_CAT(B, A),
#define VARG_DEF_LIST(T, N) BOOST_PP_REPEAT(N, VARG_FORMAT, T varg) \
                            BOOST_PP_CAT(T varg, N)
#define VARG_LIST(N) BOOST_PP_REPEAT(N, VARG_FORMAT, varg) \
                     BOOST_PP_CAT(varg, N)
#define VARG_ARRAY(N) { VARG_LIST(N) }

// Evaluates formula using the Excel parser
__declspec(dllexport) XLOPER12 WINAPI eval(XLOPER12 formula) {
    XLOPER12 result;
    Excel12(xlfEvaluate, &result, 1, formula);
    result.xltype |= xlbitXLFree;
    return result;
}

enum ArrayConsType { ArrayConsRows, ArrayConsColumns };

LPXLOPER12 array_cons(ArrayConsType array_type, VARG_DEF_LIST(LPXLOPER12, VARG_MAX)) {
    LPXLOPER12 vargs[] = VARG_ARRAY(VARG_MAX);
    int args_passed = 0;
    for(int i = 0; i < VARG_MAX; ++i, ++args_passed) {
        if (vargs[i]->xltype == xltypeMissing) {
            break;
        }
    }
    XLOPER12 list;
    list.xltype = xltypeMulti | xlbitDLLFree;
    list.val.array.lparray = new XLOPER12[args_passed];
    if (array_type == ArrayConsColumns) {
        list.val.array.rows = 1;
        list.val.array.columns = args_passed;
    } else {
        list.val.array.rows = args_passed;
        list.val.array.columns = 1;
    }
    for(int i = 0; i < args_passed; ++i) {
        list.val.array.lparray[i] = *vargs[i];
    }
    return &list;
}

__declspec(dllexport) LPXLOPER12 WINAPI r_(VARG_DEF_LIST(LPXLOPER12, VARG_MAX)) {
    return array_cons(ArrayConsRows, VARG_LIST(VARG_MAX));
}

__declspec(dllexport) LPXLOPER12 WINAPI c_(VARG_DEF_LIST(LPXLOPER12, VARG_MAX)) {
    return array_cons(ArrayConsColumns, VARG_LIST(VARG_MAX));
}

__declspec(dllexport) void WINAPI xlAutoFree12(LPXLOPER12 p) {
    // Arrays created by DLL
    if (p->xltype == (xltypeMulti | xlbitDLLFree)) {
        delete [] p->val.array.lparray;
    // Variants created by Excel
    } else if (p->xltype & xlbitXLFree) {
        Excel12(xlFree, 0, 1, p);
    }
}

__declspec(dllexport) char* ctype(LPXLOPER12 p) {
    switch (p->xltype) {
    case xltypeBigData:
        return "xltypeBigData";
    case xltypeBool:
        return "xltypeBool";
    case xltypeErr:
        return "xltypeErr";
    case xltypeFlow:
        return "xltypeFlow";
    case xltypeInt:
        return "xltypeInt";
    case xltypeMissing:
        return "xltypeMissing";
    case xltypeMulti:
        return "xltypeMulti";
    case xltypeNil:
        return "xltypeNil";
    case xltypeNum:
        return "xltypeNum";
    case xltypeRef:
        return "xltypeRef";
    case xltypeSRef:
        return "xltypeSRef";
    case xltypeStr:
        return "xltypeStr";
    }
    return "Unknown";
}