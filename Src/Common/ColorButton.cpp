/** 
 * @file  ColorButton.cpp
 *
 * @brief Implementation file for CColorButton
 *
 */

#include "stdafx.h"
#include "ColorButton.h"
#include "MergeDarkMode.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

/** 
 * @brief Default constructor setting color to black.
 */
CColorButton::CColorButton()
: m_clrFill(RGB(0, 0, 0))
{
}

/** 
 * @brief Constructor setting initial color
 * @param [in] clfFill Initial color to set.
 */
CColorButton::CColorButton(COLORREF clrFill)
: m_clrFill(clrFill)
{
}

/** 
 * @brief Draws color button
 */
void CColorButton::DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct) 
{
#if defined(USE_DARKMODELIB)
	if (DarkMode::isEnabled())
		DarkMode::paintRoundRect(lpDrawItemStruct->hDC, lpDrawItemStruct->rcItem, DarkMode::getEdgePen(), CBrush(m_clrFill));
	else
#endif
		::FillRect(lpDrawItemStruct->hDC, &lpDrawItemStruct->rcItem, CBrush(m_clrFill));
}

/** 
 * @brief Sets new button color
 * @param [in] crlFill New color
 * @param [in] bInvalidate If `true` button is invalidated (causing redraw)
 */
void CColorButton::SetColor(COLORREF clrFill, bool bInvalidate /* = true*/)
{
	m_clrFill = clrFill;
	if (bInvalidate && m_hWnd != nullptr)
		Invalidate();
}
