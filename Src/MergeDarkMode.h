/////////////////////////////////////////////////////////////////////////////
//    Dark mode for WinMerge
//    Copyright (C) 2025  ozone10
//    SPDX-License-Identifier: GPL-3.0-or-later
/////////////////////////////////////////////////////////////////////////////
/**
 * @file  MergeDarkMode.h
 *
 * @brief Declaration file for Dark mode for WinMerge
 *
 */

#pragma once

#include "DarkModeSubclass.h"

// allow only x64 and arm64 for compatibility for older OS
#if !defined(_DARKMODELIB_NOT_USED) \
	&& (defined(__x86_64__) || defined(_M_X64) \
	|| defined(__arm64__) || defined(__arm64) || defined(_M_ARM64))
#define USE_DARKMODELIB
#endif

constexpr auto MERGE_INI_NAME = L"winMergeDark"; // replace with L"" to disable ini config
constexpr DWORD CC_FLAGS_DARK = CC_RGBINIT | CC_FULLOPEN | CC_ENABLEHOOK;

namespace ATL
{
	class CImage; // from atlimage.h
};

namespace WinMergeDarkMode
{
	void InvertLightness(ATL::CImage& image);
	void SubclassAsciiArt(HWND hWnd);
}
