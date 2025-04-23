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

namespace ATL
{
	class CImage; // from atlimage.h
};

namespace WinMergeDarkMode
{
	void InvertLightness(ATL::CImage& image);
	void SubclassAsciiArt(HWND hWnd);
}
