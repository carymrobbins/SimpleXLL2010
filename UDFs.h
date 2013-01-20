#ifndef UDFS_H
#define UDFS_H

#define g_rgNumUDFs 5
#define g_rgUDFdata 11

#include <windows.h>
#include <boost/preprocessor.hpp>

#define SIG_ARG_MAX 200
#define TOKEN_FOR_PP_REPEAT(Z, A, B) B
#define REPEAT_TOKEN(N, A) BOOST_PP_REPEAT(N, TOKEN_FOR_PP_REPEAT, A)

// Function signature types:
// http://msdn.microsoft.com/en-us/library/office/bb687869.aspx

static LPWSTR g_rgUDFs[g_rgNumUDFs][g_rgUDFdata] =
{
	{
		L"eval", // Function name/ordinal
		L"QQ", // Function signature type
		L"eval", // Func name in Func wizard
		L"formula",	// Arg name in Func wizard
		L"1", // Function type
		L"SimpleXll2010", // Category in Func wizard
		L"", // Shortcut (commands only)
		L"", // Help topic
		L"Evaluates formula string and returns the result", // Func help in Func wizard
		L"Formula text to evaluate", // Arg help in Func wizard
		L""	// Arg help in Func wizard
	},
    {
		L"ctype", // Function name/ordinal
        L"CU", // Function signature type
		L"ctype", // Func name in Func wizard
		L"p",	// Arg name in Func wizard
		L"1", // Function type
		L"SimpleXll2010", // Category in Func wizard
		L"", // Shortcut (commands only)
		L"", // Help topic
		L"", // Func help in Func wizard
		L"", // Arg help in Func wizard
		L""	// Arg help in Func wizard
	},
    {
		L"r_", // Function name/ordinal
        REPEAT_TOKEN(SIG_ARG_MAX, L"Q"), // Function signature type
		L"r_", // Func name in Func wizard
		L"args...",	// Arg name in Func wizard
		L"1", // Function type
		L"SimpleXll2010", // Category in Func wizard
		L"", // Shortcut (commands only)
		L"", // Help topic
		L"", // Func help in Func wizard
		L"", // Arg help in Func wizard
		L""	// Arg help in Func wizard
	},
    {
		L"c_", // Function name/ordinal
        REPEAT_TOKEN(SIG_ARG_MAX, L"Q"), // Function signature type
		L"c_", // Func name in Func wizard
		L"args...",	// Arg name in Func wizard
		L"1", // Function type
		L"SimpleXll2010", // Category in Func wizard
		L"", // Shortcut (commands only)
		L"", // Help topic
		L"", // Func help in Func wizard
		L"", // Arg help in Func wizard
		L""	// Arg help in Func wizard
	},
};

#endif