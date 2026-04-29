// ====================== NOVATUORIA   ORIGINAL   PROJECT =====================
//
//
//                //|| //  //==//  //==//    //==//  //  //  //==//
//               // ||//  //  //  //==//    //==//  //  //  //==//
//              //  |//  //==//  //        //      //==//  //
//
//
//           NOP PUP Visual Manager for MFC (ver. 0.1 "helloWorld!")           
//
//
//         Проект: NOP PUP Visual Manager
//         Версия: ver. 0.1-alpha
//          Автор: Михайлова Александра (@novatuoria)
//         GitHub: https://github.com/novatuoria
//     Repository: https://github.com/novatuoria/NOP_PUP_VisualManager
//        License: MIT
//
//                                   2026
//
// ===================== PROPER   UNIVERSAL   PERFORMANCES ====================
//
// Примечание: 
// Перевод комментариев выполнен при поддержке ИИ Gemini.
//
// Спасибо, что решили уделить время моему проекту! :)
//
// ----------------------------------------------------------------------------
//
// English Note: 
// Technical comment translations by AI (Gemini).
//
// Thank you for taking the time to explore this project! :)
//
// ============================================================================
// 
// [ ВНИМАНИЕ !!! ATTENTION ]
// Если ваш проект использует предкомпилированные заголовки (PCH), 
// добавьте #include "pch.h" (или "stdafx.h") первой строкой в этот файл,
// ЛИБО отключите PCH для этого конкретного файла в настройках проекта 
// (C++/Precompiled Headers -> Not Using).
//
// If your project uses Precompiled Headers (PCH), 
// add #include "pch.h" (or "stdafx.h") as the first line of this file,
// OR disable PCH for this specific file in the project settings.
//
// ----------------------------------------------------------------------------
//
// [ Структура файла / File Structure ]
// Данный файл состоит из следующих регионов, которые группируют методы 
// визуального менеджера (this file is organized into the following regions, 
// grouping the visual manager's methods) :
//
//		Region InsideClass
//
//		Region Standard
//			|__ Region MainFrame
//			|__ Region AutoHide
//			|__ Region MiniFrame
//			|__ Region PropertyGrid
//			|__ Region Tabs
//			|__ Region OutlookBar
//			|__ Region TaskPane
//			|__ Region Other
//
//		Region Ribbon
//			|__ MainFrameRibbon
//			|__ TabsRBBN
//			|__ MainMenuRBBN
//			|__ OtherRBBN
//


#include "pch.h"

#include "NOP_PUP_VisualManager.h"
#include <typeinfo>

// ============================================================================
// ВНУТРЕННИЕ ФУНКЦИИ КЛАССА / INSIDE CLASS FUNCTIONS
// ----------------------------------------------------------------------------
// Жизненный цикл (конструкторы/деструкторы) и служебные методы управления.
// Обеспечивают внутреннюю логику и взаимодействие с состоянием визуального 
// менеджера.
//
// Lifecycle (constructors/destructors) and internal management methods.
// Handles internal logic and state transitions of the visual manager.
// ============================================================================
#pragma region InsideClass

IMPLEMENT_DYNCREATE(CNPUP_VisualManager, CMFCVisualManagerOfficeXP);

CNPUP_VisualManager::ThemeStyle CNPUP_VisualManager::m_currentStyle;
CNPUP_VisualManager::ColorTheme CNPUP_VisualManager::m_currentTheme;
CNPUP_VisualManager::DefaultColorTheme CNPUP_VisualManager::m_DefaultColor;
CMenuImages CNPUP_VisualManager::m_CMIUserImage;

CNPUP_VisualManager::CNPUP_VisualManager()
{
	ColorTheme::ParametrsCMenuImage prmDrawMenu = m_currentTheme.m_sParamCMenu;

	m_CMIUserImage.CleanUp();
	m_CMIUserImage.SetColor(CMenuImages::ImageWhite, prmDrawMenu.clr_UserColorImageIcon);
	m_CMIUserImage.SetColor(CMenuImages::ImageBlack, prmDrawMenu.clr_UserColorImageIcon);
	m_CMIUserImage.SetColor(CMenuImages::ImageGray, prmDrawMenu.clr_UserColorImageIcon);
	m_CMIUserImage.SetColor(CMenuImages::ImageLtGray, prmDrawMenu.clr_UserColorImageIcon);
	m_CMIUserImage.SetColor(CMenuImages::ImageDkGray, prmDrawMenu.clr_UserColorImageIcon);
	m_CMIUserImage.SetColor(CMenuImages::ImageBlack2, prmDrawMenu.clr_UserColorImageIcon);

}

CNPUP_VisualManager::~CNPUP_VisualManager()
{

}


void CNPUP_VisualManager::NVS_DrawAlphaChanel
(
	CDC * clientDC,
	BYTE alphaChanel
)
{
	BITMAP bm;
	CBitmap *pBitmap;
	pBitmap = clientDC->GetCurrentBitmap();

	pBitmap->GetObject(sizeof(BITMAP), &bm);

	if (bm.bmBitsPixel == 32)
	{
		unsigned char *pData = new unsigned char[bm.bmHeight * bm.bmWidthBytes];

		pBitmap->GetBitmapBits(bm.bmHeight * bm.bmWidthBytes, pData);

		for (int y = 0; y < bm.bmHeight; y++)
		{
			for (int x = 0; x < bm.bmWidth; x++)
			{
				pData[x * 4 + 3 + y * bm.bmWidthBytes] = alphaChanel;
			}
		}

		pBitmap->SetBitmapBits(bm.bmHeight * bm.bmWidthBytes, pData);

		delete[] pData;
	}

	clientDC->SelectObject(&pBitmap);

}

void CNPUP_VisualManager::NVS_DrawRectBorder
(
	CDC * clientDC,
	CRect rectDraw,
	COLORREF colorRGB,
	int sizeTopBorder,
	int sizeBottomBorder,
	int sizeLeftBorder,
	int sizeRightBorder
)
{
	CRgn rgnMain;
	CRgn rgnExclude;

	rgnMain.CreateRectRgn(rectDraw.left, rectDraw.top, rectDraw.right, rectDraw.bottom);
	rgnExclude.CreateRectRgn(
		rectDraw.left + sizeLeftBorder,
		rectDraw.top + sizeTopBorder,
		rectDraw.right - sizeRightBorder,
		rectDraw.bottom - sizeBottomBorder);
	rgnMain.CombineRgn(&rgnMain, &rgnExclude, RGN_DIFF);
	clientDC->SelectClipRgn(&rgnMain);
	clientDC->FillSolidRect(rectDraw, colorRGB);
	clientDC->SelectClipRgn(NULL);
	rgnMain.DeleteObject();
	rgnExclude.DeleteObject();
}

void CNPUP_VisualManager::SetColorThemeGlobal()
{
	DefaultColorTheme clrDefault = m_DefaultColor;

	/*Global*/
	ColorTheme::ColorGlobalElement &clrGlblElem = m_currentTheme.m_ColorGlobalElement;

	clrGlblElem.clr_BackgroundFrame = clrDefault.clr_ClobalBackground;
	clrGlblElem.clr_BackgroundClientFrame = clrDefault.clr_ClientFrameBackground;
	clrGlblElem.clr_DividerFrame = clrDefault.clr_ClientBorder;
	clrGlblElem.clr_BorderFrame = clrDefault.clr_Caption;

	clrGlblElem.clr_BackgroundMenu = clrDefault.clr_DisableCaption;
	clrGlblElem.clr_BackgroundLabelMenu = clrDefault.clr_Caption;
	clrGlblElem.clr_BorderMenu = clrDefault.clr_ClientBorder;
	clrGlblElem.clr_Separator = clrDefault.clr_Caption;
	clrGlblElem.clr_SelectedMenu = clrDefault.clr_ActiveCaption;
	clrGlblElem.clr_ImportantMenu = RGB(128, 128, 128);

	clrGlblElem.clr_CheckMenu = clrDefault.clr_Caption;
	clrGlblElem.clr_ResizeBar = clrDefault.clr_Caption;
	clrGlblElem.clr_ShowAllMenuBackground = clrDefault.clr_ActiveCaption;
	clrGlblElem.clr_ShowAllMenuBackground_Pressed = clrDefault.clr_DisableCaption;
	clrGlblElem.clr_ShowAllMenuBackground_Highlight = clrDefault.clr_ActiveCaption;

	clrGlblElem.clr_CheckBoxBackground = clrDefault.clr_ActiveSysButton;
	clrGlblElem.clr_CheckBoxBackground_Enable = clrDefault.clr_ActiveSysButton;
	clrGlblElem.clr_CheckBoxBackground_Disable = clrDefault.clr_DisableSysButton;
	clrGlblElem.clr_CheckBoxBackground_Highlight = clrDefault.clr_HighlightSysButton;
	clrGlblElem.clr_CheckBoxBackground_Pressed = clrDefault.clr_PressedSysButton;

	clrGlblElem.clr_CheckBoxBorder = clrDefault.clr_BorderActiveSysButton;
	clrGlblElem.clr_CheckBoxBorder_Enable = clrDefault.clr_BorderActiveSysButton;
	clrGlblElem.clr_CheckBoxBorder_Disable = clrDefault.clr_BorderDisableSysButton;
	clrGlblElem.clr_CheckBoxBorder_Highlight = clrDefault.clr_BorderHighlightSysButton;
	clrGlblElem.clr_CheckBoxBorder_Pressed = clrDefault.clr_BorderPressedSysButton;

	clrGlblElem.clr_SpinButtonsBackground_Up = clrDefault.clr_ButtonClient;
	clrGlblElem.clr_SpinButtonsBackground_Down = clrDefault.clr_ButtonClient;
	clrGlblElem.clr_SpinButtonsBackgroundPressed_Up = clrDefault.clr_PressedButtonClient;
	clrGlblElem.clr_SpinButtonsBackgroundPressed_Down = clrDefault.clr_PressedButtonClient;
	clrGlblElem.clr_SpinButtonsBackgroundHighlight_Up = clrDefault.clr_HighlightSysButton;
	clrGlblElem.clr_SpinButtonsBackgroundHighlight_Down = clrDefault.clr_HighlightSysButton;
	clrGlblElem.clr_SpinButtonsBackgroundDisable_Up = clrDefault.clr_DisableButtonClient;
	clrGlblElem.clr_SpinButtonsBackgroundDisable_Down = clrDefault.clr_DisableButtonClient;

	clrGlblElem.clr_SpinButtonsBorder_Up = clrDefault.clr_BorderButtonClient;
	clrGlblElem.clr_SpinButtonsBorder_Down = clrDefault.clr_BorderButtonClient;
	clrGlblElem.clr_SpinButtonsBorderPressed_Up = clrDefault.clr_BorderPressedButtonClient;
	clrGlblElem.clr_SpinButtonsBorderPressed_Down = clrDefault.clr_BorderPressedButtonClient;
	clrGlblElem.clr_SpinButtonsBorderHighlight_Up = clrDefault.clr_BorderHighlightSysButton;
	clrGlblElem.clr_SpinButtonsBorderHighlight_Down = clrDefault.clr_BorderHighlightSysButton;
	clrGlblElem.clr_SpinButtonsBorderDisable_Up = clrDefault.clr_BorderDisableButtonClient;
	clrGlblElem.clr_SpinButtonsBorderDisable_Down = clrDefault.clr_BorderDisableButtonClient;

	clrGlblElem.clr_ToolTipInfoBackground = clrDefault.clr_Caption;
	clrGlblElem.clr_ToolTipInfoGradient = clrDefault.clr_Caption;
	clrGlblElem.clr_ToolTipInfoBorder = clrDefault.clr_ClientBorder;

	clrGlblElem.clr_CommandListBackground = clrDefault.clr_Caption;
	clrGlblElem.clr_CommandListBackground_Selected = clrDefault.clr_ActiveCaption;
	clrGlblElem.clr_CommandListBackground_Highlight = clrDefault.clr_HighlightSysButton;

	clrGlblElem.clr_TextFrame = clrDefault.clr_TextActive;
	clrGlblElem.clr_TextLabelMenu = clrDefault.clr_TextActive;
	clrGlblElem.clr_TextMenu = clrDefault.clr_Text;
	clrGlblElem.clr_TextMenu_Selected = clrDefault.clr_TextActive;
	clrGlblElem.clr_TextMenu_Pressed = clrDefault.clr_TextPressed;
	clrGlblElem.clr_TextMenu_Disable = clrDefault.clr_TextDisable;
	clrGlblElem.clr_TextMenu_Highlight = clrDefault.clr_TextActive;
	clrGlblElem.clr_ToolTipInfoText = clrDefault.clr_TextActive;
	clrGlblElem.clr_TextCommandList = clrDefault.clr_TextActive;
	clrGlblElem.clr_TextCommandList_Selected = clrDefault.clr_TextActive;
	clrGlblElem.clr_TextCommandList_Highlight = clrDefault.clr_TextActive;

	/*Miniframe*/
	ColorTheme::ColorMiniFrame &minFrameClr = m_currentTheme.m_ColorMiniFrame;

	minFrameClr.clr_BackgroundMiniFrame = clrDefault.clr_Caption;
	minFrameClr.clr_BackgroundMiniFrame_CaptionActive = clrDefault.clr_ActiveCaption;
	minFrameClr.clr_BackgroundMiniFrame_CaptionDisable = clrDefault.clr_Caption;
	minFrameClr.clr_BackgroundMiniFrame_BorderActive = clrDefault.clr_Caption;
	minFrameClr.clr_BackgroundMiniFrame_BorderDisable = clrDefault.clr_Caption;
	minFrameClr.clr_BackgroundMiniFrame_Client = clrDefault.clr_ClientFrameBackground;

	minFrameClr.clr_BorderMiniFrame_Client = clrDefault.clr_Caption;

	minFrameClr.clr_ButtonMiniFrame = clrDefault.clr_Caption;
	minFrameClr.clr_ButtonMiniFrame_Active = clrDefault.clr_HighlightSysButton;
	minFrameClr.clr_ButtonMiniFrame_Disable = clrDefault.clr_Caption;
	minFrameClr.clr_ButtonMiniFrame_Highlight = clrDefault.clr_ActiveCaption;

	minFrameClr.clr_ButtonBorderMiniFrame_Active = clrDefault.clr_HighlightSysButton;
	minFrameClr.clr_ButtonBorderMiniFrame_Disable = clrDefault.clr_Caption;
	minFrameClr.clr_ButtonBorderMiniFrame_Highlight = clrDefault.clr_ActiveCaption;

	minFrameClr.clr_TextCaptionMiniFrame = clrDefault.clr_TextDisable;
	minFrameClr.clr_TextCaptionMiniFrame_Active = clrDefault.clr_TextActive;
	minFrameClr.clr_TextCaptionMiniFrame_Disable = clrDefault.clr_TextDisable;

	minFrameClr.clr_ButtonsElement = clrDefault.clr_SysButton;
	minFrameClr.clr_ButtonsElement_Active = clrDefault.clr_SysButton;
	minFrameClr.clr_ButtonsElement_Disable = clrDefault.clr_DisableSysButton;
	minFrameClr.clr_ButtonsElement_Dropped = clrDefault.clr_ActiveSysButton;
	minFrameClr.clr_ButtonsElement_NoDropped = clrDefault.clr_ActiveSysButton;
	minFrameClr.clr_ButtonsElement_Highlight = clrDefault.clr_HighlightSysButton;

	minFrameClr.clr_ButtonsBorderElement_Active = clrDefault.clr_BorderSysButton;
	minFrameClr.clr_ButtonsBorderElement_Disable = clrDefault.clr_BorderDisableSysButton;
	minFrameClr.clr_ButtonsBorderElement_Dropped = clrDefault.clr_BorderActiveSysButton;
	minFrameClr.clr_ButtonsBorderElement_NoDropped = clrDefault.clr_BorderActiveSysButton;
	minFrameClr.clr_ButtonsBorderElement_Highlight = clrDefault.clr_BorderHighlightSysButton;

	/*Tabs*/
	ColorTheme::ColorTabs &tabClr = m_currentTheme.m_ColorTabs;

	tabClr.clr_BackgroundFrameDockMDI = clrDefault.clr_ClientBorder;
	tabClr.clr_BackgroundFrameDockHide = clrDefault.clr_ClientBorder;

	tabClr.clr_BackgroundTabsHide = clrDefault.clr_Caption;
	tabClr.clr_BackgroundTabsHide_Active = clrDefault.clr_ActiveCaption;
	tabClr.clr_BackgroundTabsHide_Disable = clrDefault.clr_Caption;
	tabClr.clr_BackgroundTabsHide_Highlight = clrDefault.clr_ActiveCaption;

	tabClr.clr_BackgroundTabsMDI = clrDefault.clr_Caption;
	tabClr.clr_BackgroundTabsMDI_Active = clrDefault.clr_ActiveCaption;
	tabClr.clr_BackgroundTabsMDI_Disable = clrDefault.clr_Caption;
	tabClr.clr_BackgroundTabsMDI_Highlight = clrDefault.clr_ActiveCaption;

	tabClr.clr_BorderTabsMDI = clrDefault.clr_Caption;
	tabClr.clr_BorderTabsMDI_Active = clrDefault.clr_ActiveCaption;
	tabClr.clr_BorderTabsMDI_Disable = clrDefault.clr_ClobalBackground;
	tabClr.clr_BorderTabsMDI_Highlight = clrDefault.clr_ActiveCaption;

	tabClr.clr_BorderTabsHide = clrDefault.clr_Caption;
	tabClr.clr_BorderTabsHide_Active = clrDefault.clr_ActiveCaption;
	tabClr.clr_BorderTabsHide_Disable = clrDefault.clr_ClobalBackground;
	tabClr.clr_BorderTabsHide_Highlight = clrDefault.clr_ActiveCaption;

	tabClr.clr_ButtonTabsMDI = clrDefault.clr_Caption;
	tabClr.clr_ButtonTabsMDI_Active = clrDefault.clr_ActiveCaption;
	tabClr.clr_ButtonTabsMDI_Disable = clrDefault.clr_Caption;
	tabClr.clr_ButtonTabsMDI_Highlight = clrDefault.clr_HighlightSysButton;

	tabClr.clr_BorderButtonTabsMDI = clrDefault.clr_Caption;
	tabClr.clr_BorderButtonTabsMDI_Active = clrDefault.clr_ActiveCaption;
	tabClr.clr_BorderButtonTabsMDI_Disable = clrDefault.clr_Caption;
	tabClr.clr_BorderButtonTabsMDI_Highlight = clrDefault.clr_HighlightSysButton;

	tabClr.clr_TextTabsMDI = clrDefault.clr_Text;
	tabClr.clr_TextTabsMDI_Active = clrDefault.clr_TextActive;
	tabClr.clr_TextTabsMDI_Disable = clrDefault.clr_TextDisable;
	tabClr.clr_TextTabsMDI_Highlight = clrDefault.clr_TextActive;

	tabClr.clr_TextTabsHide = clrDefault.clr_Text;
	tabClr.clr_TextTabsHide_Active = clrDefault.clr_TextActive;
	tabClr.clr_TextTabsHide_Disable = clrDefault.clr_TextDisable;
	tabClr.clr_TextTabsHide_Highlight = clrDefault.clr_TextActive;

	/*PropertyGrid*/
	ColorTheme::ColorPropertyGrid &propGridColor = m_currentTheme.m_ColorPropertyGrid;

	propGridColor.clr_PropGrid_Background = clrDefault.clr_ClientFrameBackground;
	propGridColor.clr_PropGrid_GroupBackground = clrDefault.clr_Caption;
	propGridColor.clr_PropGrid_DescriptionBackground = clrDefault.clr_Caption;
	propGridColor.clr_PropGrid_Line = clrDefault.clr_ClientBorder;

	propGridColor.clr_PropGrid_Text = clrDefault.clr_Text;
	propGridColor.clr_PropGrid_GroupText = clrDefault.clr_Text;
	propGridColor.clr_PropGrid_DescriptionText = clrDefault.clr_Text;

	/*ThreeView*/
	ColorTheme::ColorThreeView & clrThreeView = m_currentTheme.m_ColorThreeView;

	clrThreeView.clr_ThreeViewBackground = clrDefault.clr_ClientFrameBackground;
	clrThreeView.clr_ThreeViewLine = clrDefault.clr_ClientBorder;

	clrThreeView.clr_ThreeView_Text = clrDefault.clr_Text;

	/*Toolbars*/
	ColorTheme::ColorToolbars &toolBarClr = m_currentTheme.m_ColorToolbars;

	toolBarClr.clr_BackgroundToolbar = clrDefault.clr_Caption;
	toolBarClr.clr_BorderToolbar = clrDefault.clr_Caption;

	toolBarClr.clr_BackgroundIconToolbar = clrDefault.clr_Caption;
	toolBarClr.clr_BackgroundIconToolbar_Pressed = clrDefault.clr_DisableCaption;
	toolBarClr.clr_BackgroundIconToolbar_Highlight = clrDefault.clr_ActiveCaption;

	toolBarClr.clr_BorderIconToolbar = clrDefault.clr_Caption;
	toolBarClr.clr_BorderIconToolbar_Pressed = clrDefault.clr_DisableCaption;
	toolBarClr.clr_BorderIconToolbar_Highlight = clrDefault.clr_ActiveCaption;

	toolBarClr.clr_ToolbarGripper = clrDefault.clr_Caption;

	/*PaneCaption*/
	ColorTheme::ColorPaneCaptionsAndButton &paneCaptionClr = m_currentTheme.m_ColorPaneCaptionsAndButton;

	paneCaptionClr.clr_PaneCaption_BackgroundFrame = clrDefault.clr_ClobalBackground;
	paneCaptionClr.clr_PaneCaption_BackgroundCaption = clrDefault.clr_ActiveCaption;
	paneCaptionClr.clr_PaneCaption_Border = clrDefault.clr_Caption;

	paneCaptionClr.clr_PaneCaption_ButtonsBackground = clrDefault.clr_ButtonClient;
	paneCaptionClr.clr_PaneCaption_ButtonsBackground_Disable = clrDefault.clr_DisableButtonClient;
	paneCaptionClr.clr_PaneCaption_ButtonsBackground_Pressed = clrDefault.clr_PressedButtonClient;
	paneCaptionClr.clr_PaneCaption_ButtonsBackground_Highlight = clrDefault.clr_ActiveButtonClient;

	paneCaptionClr.clr_PaneCaption_ButtonsBorder = clrDefault.clr_BorderButtonClient;
	paneCaptionClr.clr_PaneCaption_ButtonsBorder_Disable = clrDefault.clr_BorderDisableButtonClient;
	paneCaptionClr.clr_PaneCaption_ButtonsBorder_Pressed = clrDefault.clr_BorderPressedButtonClient;
	paneCaptionClr.clr_PaneCaption_ButtonsBorder_Highlight = clrDefault.clr_BorderActiveButtonClient;

	paneCaptionClr.clr_PaneCaption_Text = clrDefault.clr_Text;

	paneCaptionClr.clr_PaneCaption_ButtonsText = clrDefault.clr_Text;
	paneCaptionClr.clr_PaneCaption_ButtonsText_Disable = clrDefault.clr_TextDisable;
	paneCaptionClr.clr_PaneCaption_ButtonsText_Pressed = clrDefault.clr_TextPressed;
	paneCaptionClr.clr_PaneCaption_ButtonsText_Highlight = clrDefault.clr_TextActive;

	/*OutlookBar*/
	ColorTheme::ColorOutlookBar &outlookBarClr = m_currentTheme.m_ColorOutlookBar;
	outlookBarClr.clr_OutlookBarCaptionBackground = RGB(0, 16, 64);
	outlookBarClr.clr_OutlookBarSeparator = RGB(0, 64, 64);

	outlookBarClr.clr_OutlookBarButtonBackground = RGB(0, 64, 80);
	outlookBarClr.clr_OutlookBarButtonBackground_Pressed = RGB(0, 0, 32);
	outlookBarClr.clr_OutlookBarButtonBackground_Highlight = RGB(0, 128, 128);

	outlookBarClr.clr_OutlookBarButtonBorder = RGB(0, 32, 32);
	outlookBarClr.clr_OutlookBarButtonBorder_Pressed = RGB(0, 255, 255);
	outlookBarClr.clr_OutlookBarButtonBorder_Highlight = RGB(128, 128, 0);

	outlookBarClr.clr_OutlookBarCaptionText = RGB(224, 224, 224);
	outlookBarClr.clr_OutlookBarButtonText = RGB(224, 224, 224);
	outlookBarClr.clr_OutlookBarButtonText_Pressed = RGB(128, 128, 128);
	outlookBarClr.clr_OutlookBarButtonText_Highlight = RGB(255, 255, 255);

	/*FrameRibbon*/
	ColorTheme::ColorRibbonBackground &clrFrameRibbon = m_currentTheme.m_ColorRibbonBackground;

	clrFrameRibbon.clr_RibbonBackground = clrDefault.clr_ClobalBackground;

	clrFrameRibbon.clr_RibbonCategoryCaption = clrDefault.clr_RibbonFrame;
	clrFrameRibbon.clr_RibbonCategoryCaption_Red = clrDefault.clr_RibbonFrame_Red;
	clrFrameRibbon.clr_RibbonCategoryCaption_Orange = clrDefault.clr_RibbonFrame_Orange;
	clrFrameRibbon.clr_RibbonCategoryCaption_Yellow = clrDefault.clr_RibbonFrame_Yellow;
	clrFrameRibbon.clr_RibbonCategoryCaption_Green = clrDefault.clr_RibbonFrame_Green;
	clrFrameRibbon.clr_RibbonCategoryCaption_Blue = clrDefault.clr_RibbonFrame_Blue;
	clrFrameRibbon.clr_RibbonCategoryCaption_Indigo = clrDefault.clr_RibbonFrame_Indigo;
	clrFrameRibbon.clr_RibbonCategoryCaption_Violet = clrDefault.clr_RibbonFrame_Violet;

	clrFrameRibbon.clr_RibbonCategoryCaption_Highlight = clrDefault.clr_RibbonFrame;
	clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Red = clrDefault.clr_RibbonFrame_Red;
	clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Orange = clrDefault.clr_RibbonFrame_Orange;
	clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Yellow = clrDefault.clr_RibbonFrame_Yellow;
	clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Green = clrDefault.clr_RibbonFrame_Green;
	clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Blue = clrDefault.clr_RibbonFrame_Blue;
	clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Indigo = clrDefault.clr_RibbonFrame_Indigo;
	clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Violet = clrDefault.clr_RibbonFrame_Violet;

	clrFrameRibbon.clr_RibbonCategoryCaption_Active = clrDefault.clr_RibbonFrame;
	clrFrameRibbon.clr_RibbonCategoryCaption_Active_Red = clrDefault.clr_RibbonFrame_Red;
	clrFrameRibbon.clr_RibbonCategoryCaption_Active_Orange = clrDefault.clr_RibbonFrame_Orange;
	clrFrameRibbon.clr_RibbonCategoryCaption_Active_Yellow = clrDefault.clr_RibbonFrame_Yellow;
	clrFrameRibbon.clr_RibbonCategoryCaption_Active_Green = clrDefault.clr_RibbonFrame_Green;
	clrFrameRibbon.clr_RibbonCategoryCaption_Active_Blue = clrDefault.clr_RibbonFrame_Blue;
	clrFrameRibbon.clr_RibbonCategoryCaption_Active_Indigo = clrDefault.clr_RibbonFrame_Indigo;
	clrFrameRibbon.clr_RibbonCategoryCaption_Active_Violet = clrDefault.clr_RibbonFrame_Violet;

	clrFrameRibbon.clr_RibbonCategoryCaptionBorder = clrDefault.clr_RibbonFrame;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Red = clrDefault.clr_RibbonFrame_Red;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Orange = clrDefault.clr_RibbonFrame_Orange;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Yellow = clrDefault.clr_RibbonFrame_Yellow;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Green = clrDefault.clr_RibbonFrame_Green;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Blue = clrDefault.clr_RibbonFrame_Blue;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Indigo = clrDefault.clr_RibbonFrame_Indigo;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Violet = clrDefault.clr_RibbonFrame_Violet;

	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight = clrDefault.clr_RibbonFrame;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Red = clrDefault.clr_RibbonFrame_Red;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Orange = clrDefault.clr_RibbonFrame_Orange;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Yellow = clrDefault.clr_RibbonFrame_Yellow;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Green = clrDefault.clr_RibbonFrame_Green;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Blue = clrDefault.clr_RibbonFrame_Blue;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Indigo = clrDefault.clr_RibbonFrame_Indigo;
	clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Violet = clrDefault.clr_RibbonFrame_Violet;

	/*MainMenuRibbon*/
	ColorTheme::ColorMainMenuRibbon &clrMainMenuRibbon = m_currentTheme.m_ColorMainMenuRibbon;

	clrMainMenuRibbon.clr_RibbonMainMenuApplicationButtonBackground = clrDefault.clr_SysButton;
	clrMainMenuRibbon.clr_RibbonMainMenuApplicationButtonBackground_Disabled = clrDefault.clr_DisableSysButton;
	clrMainMenuRibbon.clr_RibbonMainMenuApplicationButtonBackground_Highlight = clrDefault.clr_ActiveSysButton;
	clrMainMenuRibbon.clr_RibbonMainMenuApplicationButtonBackground_Pressed = clrDefault.clr_PressedSysButton;

	clrMainMenuRibbon.clr_RibbonMainMenuMainPanelBackground = clrDefault.clr_Caption;

	clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground = clrDefault.clr_SysButton;
	clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground_Disabled = clrDefault.clr_DisableSysButton;
	clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground_Highlight = clrDefault.clr_ActiveSysButton;
	clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground_Pressed = clrDefault.clr_PressedSysButton;

	clrMainMenuRibbon.clr_RibbonMainMenuButtonBorder = clrDefault.clr_BorderSysButton;
	clrMainMenuRibbon.clr_RibbonMainMenuButtonBorder_Disabled = clrDefault.clr_BorderDisableSysButton;
	clrMainMenuRibbon.clr_RibbonMainMenuButtonBorder_Highlight = clrDefault.clr_BorderActiveSysButton;
	clrMainMenuRibbon.clr_RibbonMainMenuButtonBorder_Pressed = clrDefault.clr_BorderPressedSysButton;

	clrMainMenuRibbon.clr_RibbonMainMenuButtonText = clrDefault.clr_Text;
	clrMainMenuRibbon.clr_RibbonMainMenuButtonText_Disabled = clrDefault.clr_TextDisable;
	clrMainMenuRibbon.clr_RibbonMainMenuButtonText_Highlight = clrDefault.clr_TextActive;
	clrMainMenuRibbon.clr_RibbonMainMenuButtonText_Pressed = clrDefault.clr_TextPressed;

	/*RibbonTabs*/
	ColorTheme::ColorRibbonTabs &clrRibbonTabs = m_currentTheme.m_ColorRibbonTabs;

	clrRibbonTabs.clr_RibbonTabs = clrDefault.clr_ClobalBackground;
	clrRibbonTabs.clr_RibbonTabs_Red = clrDefault.clr_ClobalBackground;
	clrRibbonTabs.clr_RibbonTabs_Orange = clrDefault.clr_ClobalBackground;
	clrRibbonTabs.clr_RibbonTabs_Yellow = clrDefault.clr_ClobalBackground;
	clrRibbonTabs.clr_RibbonTabs_Green = clrDefault.clr_ClobalBackground;
	clrRibbonTabs.clr_RibbonTabs_Blue = clrDefault.clr_ClobalBackground;
	clrRibbonTabs.clr_RibbonTabs_Indigo = clrDefault.clr_ClobalBackground;
	clrRibbonTabs.clr_RibbonTabs_Violet = clrDefault.clr_ClobalBackground;

	clrRibbonTabs.clr_RibbonTabs_Active = clrDefault.clr_RibbonFrame;
	clrRibbonTabs.clr_RibbonTabs_Active_Red = clrDefault.clr_RibbonFrame_Red;
	clrRibbonTabs.clr_RibbonTabs_Active_Orange = clrDefault.clr_RibbonFrame_Orange;
	clrRibbonTabs.clr_RibbonTabs_Active_Yellow = clrDefault.clr_RibbonFrame_Yellow;
	clrRibbonTabs.clr_RibbonTabs_Active_Green = clrDefault.clr_RibbonFrame_Green;
	clrRibbonTabs.clr_RibbonTabs_Active_Blue = clrDefault.clr_RibbonFrame_Blue;
	clrRibbonTabs.clr_RibbonTabs_Active_Indigo = clrDefault.clr_RibbonFrame_Indigo;
	clrRibbonTabs.clr_RibbonTabs_Active_Violet = clrDefault.clr_RibbonFrame_Violet;

	clrRibbonTabs.clr_RibbonTabs_Highlight = clrDefault.clr_RibbonFrame;
	clrRibbonTabs.clr_RibbonTabs_Highlight_Red = clrDefault.clr_RibbonFrame_Red;
	clrRibbonTabs.clr_RibbonTabs_Highlight_Orange = clrDefault.clr_RibbonFrame_Orange;
	clrRibbonTabs.clr_RibbonTabs_Highlight_Yellow = clrDefault.clr_RibbonFrame_Yellow;
	clrRibbonTabs.clr_RibbonTabs_Highlight_Green = clrDefault.clr_RibbonFrame_Green;
	clrRibbonTabs.clr_RibbonTabs_Highlight_Blue = clrDefault.clr_RibbonFrame_Blue;
	clrRibbonTabs.clr_RibbonTabs_Highlight_Indigo = clrDefault.clr_RibbonFrame_Indigo;
	clrRibbonTabs.clr_RibbonTabs_Highlight_Violet = clrDefault.clr_RibbonFrame_Violet;

	clrRibbonTabs.clr_RibbonTabsBorder = clrDefault.clr_RibbonFrame_Border;
	clrRibbonTabs.clr_RibbonTabsBorder_Red = clrDefault.clr_RibbonFrame_Border_Red;
	clrRibbonTabs.clr_RibbonTabsBorder_Orange = clrDefault.clr_RibbonFrame_Border_Orange;
	clrRibbonTabs.clr_RibbonTabsBorder_Yellow = clrDefault.clr_RibbonFrame_Border_Yellow;
	clrRibbonTabs.clr_RibbonTabsBorder_Green = clrDefault.clr_RibbonFrame_Border_Green;
	clrRibbonTabs.clr_RibbonTabsBorder_Blue = clrDefault.clr_RibbonFrame_Border_Blue;
	clrRibbonTabs.clr_RibbonTabsBorder_Indigo = clrDefault.clr_RibbonFrame_Border_Indigo;
	clrRibbonTabs.clr_RibbonTabsBorder_Violet = clrDefault.clr_RibbonFrame_Border_Violet;

	clrRibbonTabs.clr_RibbonTabsBorder_Active = clrDefault.clr_RibbonFrame_Border;
	clrRibbonTabs.clr_RibbonTabsBorder_Active_Red = clrDefault.clr_RibbonFrame_Border_Red;
	clrRibbonTabs.clr_RibbonTabsBorder_Active_Orange = clrDefault.clr_RibbonFrame_Border_Orange;
	clrRibbonTabs.clr_RibbonTabsBorder_Active_Yellow = clrDefault.clr_RibbonFrame_Border_Yellow;
	clrRibbonTabs.clr_RibbonTabsBorder_Active_Green = clrDefault.clr_RibbonFrame_Border_Green;
	clrRibbonTabs.clr_RibbonTabsBorder_Active_Blue = clrDefault.clr_RibbonFrame_Border_Blue;
	clrRibbonTabs.clr_RibbonTabsBorder_Active_Indigo = clrDefault.clr_RibbonFrame_Border_Indigo;
	clrRibbonTabs.clr_RibbonTabsBorder_Active_Violet = clrDefault.clr_RibbonFrame_Border_Violet;

	clrRibbonTabs.clr_RibbonTabsBorder_Highlight = clrDefault.clr_RibbonFrame_Border;
	clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Red = clrDefault.clr_RibbonFrame_Border_Red;
	clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Orange = clrDefault.clr_RibbonFrame_Border_Orange;
	clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Yellow = clrDefault.clr_RibbonFrame_Border_Yellow;
	clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Green = clrDefault.clr_RibbonFrame_Border_Green;
	clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Blue = clrDefault.clr_RibbonFrame_Border_Blue;
	clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Indigo = clrDefault.clr_RibbonFrame_Border_Indigo;
	clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Violet = clrDefault.clr_RibbonFrame_Border_Violet;

	clrRibbonTabs.clr_RibbonTabs_Text = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Red = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Orange = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Yellow = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Green = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Blue = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Indigo = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Violet = clrDefault.clr_Text;

	clrRibbonTabs.clr_RibbonTabs_Text_Active = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Active_Red = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Active_Orange = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Active_Yellow = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Active_Green = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Active_Blue = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Active_Indigo = clrDefault.clr_Text;
	clrRibbonTabs.clr_RibbonTabs_Text_Active_Violet = clrDefault.clr_Text;

	clrRibbonTabs.clr_RibbonTabs_Text_Highlight = clrDefault.clr_TextActive;
	clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Red = clrDefault.clr_TextActive;
	clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Orange = clrDefault.clr_TextActive;
	clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Yellow = clrDefault.clr_TextActive;
	clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Green = clrDefault.clr_TextActive;
	clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Blue = clrDefault.clr_TextActive;
	clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Indigo = clrDefault.clr_TextActive;
	clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Violet = clrDefault.clr_TextActive;

	/*RibbonCategory*/
	ColorTheme::ColorRibbonCategory &clrRibbonCategory = m_currentTheme.m_ColorRibbonCategory;

	clrRibbonCategory.clr_RibbonCategoryBackground = clrDefault.clr_ClientFrameBackground;
	clrRibbonCategory.clr_RibbonCategoryBackground_Red = clrDefault.clr_ClientFrameBackground;
	clrRibbonCategory.clr_RibbonCategoryBackground_Orange = clrDefault.clr_ClientFrameBackground;
	clrRibbonCategory.clr_RibbonCategoryBackground_Yellow = clrDefault.clr_ClientFrameBackground;
	clrRibbonCategory.clr_RibbonCategoryBackground_Green = clrDefault.clr_ClientFrameBackground;
	clrRibbonCategory.clr_RibbonCategoryBackground_Blue = clrDefault.clr_ClientFrameBackground;
	clrRibbonCategory.clr_RibbonCategoryBackground_Indigo = clrDefault.clr_ClientFrameBackground;
	clrRibbonCategory.clr_RibbonCategoryBackground_Violet = clrDefault.clr_ClientFrameBackground;

	clrRibbonCategory.clr_RibbonCategoryBorder = clrDefault.clr_RibbonFrame;
	clrRibbonCategory.clr_RibbonCategoryBorder_Red = clrDefault.clr_RibbonFrame_Red;
	clrRibbonCategory.clr_RibbonCategoryBorder_Orange = clrDefault.clr_RibbonFrame_Orange;
	clrRibbonCategory.clr_RibbonCategoryBorder_Yellow = clrDefault.clr_RibbonFrame_Yellow;
	clrRibbonCategory.clr_RibbonCategoryBorder_Green = clrDefault.clr_RibbonFrame_Green;
	clrRibbonCategory.clr_RibbonCategoryBorder_Blue = clrDefault.clr_RibbonFrame_Blue;
	clrRibbonCategory.clr_RibbonCategoryBorder_Indigo = clrDefault.clr_RibbonFrame_Indigo;
	clrRibbonCategory.clr_RibbonCategoryBorder_Violet = clrDefault.clr_RibbonFrame_Violet;

	/*RibbonPanel*/
	ColorTheme::ColorRibbonPanel &clrRibbonPanel = m_currentTheme.m_ColorRibbonPanel;

	clrRibbonPanel.clr_RibbonPanelTabBackground = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Red = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Orange = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Yellow = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Green = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Blue = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Indigo = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Violet = clrDefault.clr_ClientFrameBackground;

	clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Red = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Orange = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Yellow = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Green = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Blue = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Indigo = clrDefault.clr_ClientFrameBackground;
	clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Violet = clrDefault.clr_ClientFrameBackground;

	clrRibbonPanel.clr_RibbonPanelTabLabelBackground = clrDefault.clr_Caption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Red = clrDefault.clr_Caption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Orange = clrDefault.clr_Caption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Yellow = clrDefault.clr_Caption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Green = clrDefault.clr_Caption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Blue = clrDefault.clr_Caption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Indigo = clrDefault.clr_Caption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Violet = clrDefault.clr_Caption;

	clrRibbonPanel.clr_RibbonPanelTabLabelBorder = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Red = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Orange = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Yellow = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Green = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Blue = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Indigo = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Violet = clrDefault.clr_ActiveCaption;

	clrRibbonPanel.clr_RibbonPanelTabBorder = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Red = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Orange = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Yellow = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Green = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Blue = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Indigo = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Violet = clrDefault.clr_ActiveCaption;

	clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Red = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Orange = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Yellow = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Green = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Blue = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Indigo = clrDefault.clr_ActiveCaption;
	clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Violet = clrDefault.clr_ActiveCaption;

	clrRibbonPanel.clr_RibbonPanelTabLabelText = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Red = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Orange = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Yellow = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Green = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Blue = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Indigo = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Violet = clrDefault.clr_Text;

	clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Red = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Orange = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Yellow = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Green = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Blue = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Indigo = clrDefault.clr_Text;
	clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Violet = clrDefault.clr_Text;

	/*RibbonButton*/
	ColorTheme::ColorRibbonButton &clrRibbonButton = m_currentTheme.m_ColorRibbonButton;

	clrRibbonButton.clr_RibbonButtonBackground = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Red = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Orange = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Yellow = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Green = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Blue = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Indigo = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Violet = clrDefault.clr_ButtonClient;

	clrRibbonButton.clr_RibbonButtonBackground_Disable = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Disable_Red = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Disable_Orange = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Disable_Yellow = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Disable_Green = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Disable_Blue = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Disable_Indigo = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Disable_Violet = clrDefault.clr_DisableButtonClient;

	clrRibbonButton.clr_RibbonButtonBackground_Pressed = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Pressed_Red = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Pressed_Orange = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Pressed_Yellow = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Pressed_Green = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Pressed_Blue = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Pressed_Indigo = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Pressed_Violet = clrDefault.clr_PressedButtonClient;

	clrRibbonButton.clr_RibbonButtonBackground_Highlight = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Highlight_Red = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Highlight_Orange = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Highlight_Yellow = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Highlight_Green = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Highlight_Blue = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Highlight_Indigo = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBackground_Highlight_Violet = clrDefault.clr_ActiveButtonClient;

	clrRibbonButton.clr_RibbonButtonBorder = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Red = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Orange = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Yellow = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Green = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Blue = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Indigo = clrDefault.clr_ButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Violet = clrDefault.clr_ButtonClient;

	clrRibbonButton.clr_RibbonButtonBorder_Disable = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Disable_Red = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Disable_Orange = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Disable_Yellow = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Disable_Green = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Disable_Blue = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Disable_Indigo = clrDefault.clr_DisableButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Disable_Violet = clrDefault.clr_DisableButtonClient;

	clrRibbonButton.clr_RibbonButtonBorder_Pressed = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Pressed_Red = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Pressed_Orange = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Pressed_Yellow = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Pressed_Green = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Pressed_Blue = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Pressed_Indigo = clrDefault.clr_PressedButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Pressed_Violet = clrDefault.clr_PressedButtonClient;

	clrRibbonButton.clr_RibbonButtonBorder_Highlight = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Highlight_Red = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Highlight_Orange = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Highlight_Yellow = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Highlight_Green = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Highlight_Blue = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Highlight_Indigo = clrDefault.clr_ActiveButtonClient;
	clrRibbonButton.clr_RibbonButtonBorder_Highlight_Violet = clrDefault.clr_ActiveButtonClient;

	clrRibbonButton.clr_RibbonButtonText = clrDefault.clr_Text;
	clrRibbonButton.clr_RibbonButtonText_Red = clrDefault.clr_Text;
	clrRibbonButton.clr_RibbonButtonText_Orange = clrDefault.clr_Text;
	clrRibbonButton.clr_RibbonButtonText_Yellow = clrDefault.clr_Text;
	clrRibbonButton.clr_RibbonButtonText_Green = clrDefault.clr_Text;
	clrRibbonButton.clr_RibbonButtonText_Blue = clrDefault.clr_Text;
	clrRibbonButton.clr_RibbonButtonText_Indigo = clrDefault.clr_Text;
	clrRibbonButton.clr_RibbonButtonText_Violet = clrDefault.clr_Text;

	clrRibbonButton.clr_RibbonButtonText_Disable = clrDefault.clr_TextDisable;
	clrRibbonButton.clr_RibbonButtonText_Disable_Red = clrDefault.clr_TextDisable;
	clrRibbonButton.clr_RibbonButtonText_Disable_Orange = clrDefault.clr_TextDisable;
	clrRibbonButton.clr_RibbonButtonText_Disable_Yellow = clrDefault.clr_TextDisable;
	clrRibbonButton.clr_RibbonButtonText_Disable_Green = clrDefault.clr_TextDisable;
	clrRibbonButton.clr_RibbonButtonText_Disable_Blue = clrDefault.clr_TextDisable;
	clrRibbonButton.clr_RibbonButtonText_Disable_Indigo = clrDefault.clr_TextDisable;
	clrRibbonButton.clr_RibbonButtonText_Disable_Violet = clrDefault.clr_TextDisable;

	clrRibbonButton.clr_RibbonButtonText_Pressed = clrDefault.clr_TextPressed;
	clrRibbonButton.clr_RibbonButtonText_Pressed_Red = clrDefault.clr_TextPressed;
	clrRibbonButton.clr_RibbonButtonText_Pressed_Orange = clrDefault.clr_TextPressed;
	clrRibbonButton.clr_RibbonButtonText_Pressed_Yellow = clrDefault.clr_TextPressed;
	clrRibbonButton.clr_RibbonButtonText_Pressed_Green = clrDefault.clr_TextPressed;
	clrRibbonButton.clr_RibbonButtonText_Pressed_Blue = clrDefault.clr_TextPressed;
	clrRibbonButton.clr_RibbonButtonText_Pressed_Indigo = clrDefault.clr_TextPressed;
	clrRibbonButton.clr_RibbonButtonText_Pressed_Violet = clrDefault.clr_TextPressed;

	clrRibbonButton.clr_RibbonButtonText_Highlight = clrDefault.clr_TextActive;
	clrRibbonButton.clr_RibbonButtonText_Highlight_Red = clrDefault.clr_TextActive;
	clrRibbonButton.clr_RibbonButtonText_Highlight_Orange = clrDefault.clr_TextActive;
	clrRibbonButton.clr_RibbonButtonText_Highlight_Yellow = clrDefault.clr_TextActive;
	clrRibbonButton.clr_RibbonButtonText_Highlight_Green = clrDefault.clr_TextActive;
	clrRibbonButton.clr_RibbonButtonText_Highlight_Blue = clrDefault.clr_TextActive;
	clrRibbonButton.clr_RibbonButtonText_Highlight_Indigo = clrDefault.clr_TextActive;
	clrRibbonButton.clr_RibbonButtonText_Highlight_Violet = clrDefault.clr_TextActive;

	/*ColorRibbonSystemButton*/
	ColorTheme::ColorRibbonSystemAndQuickButton &clrSysQuickButton = m_currentTheme.m_ColorRibbonSystemAndQuickButton;

	clrSysQuickButton.clr_RibbonQuick_Background = RGB(255, 0, 0);

	clrSysQuickButton.clr_RibbonQuickButton_Background = RGB(255, 0, 0);
	clrSysQuickButton.clr_RibbonQuickButton_Background_Disable = RGB(255, 0, 0);
	clrSysQuickButton.clr_RibbonQuickButton_Background_Highlight = RGB(255, 0, 0);
	clrSysQuickButton.clr_RibbonQuickButton_Background_Pressed = RGB(255, 0, 0);

	clrSysQuickButton.clr_RibbonQuickButton_Border = RGB(255, 0, 0);
	clrSysQuickButton.clr_RibbonQuickButton_Border_Disable = RGB(255, 0, 0);
	clrSysQuickButton.clr_RibbonQuickButton_Border_Highlight = RGB(255, 0, 0);
	clrSysQuickButton.clr_RibbonQuickButton_Border_Pressed = RGB(255, 0, 0);

	clrSysQuickButton.clr_RibbonSystemButton_Background = clrDefault.clr_SysButton;
	clrSysQuickButton.clr_RibbonSystemButton_Background_Disable = clrDefault.clr_DisableSysButton;
	clrSysQuickButton.clr_RibbonSystemButton_Background_Highlight = clrDefault.clr_HighlightSysButton;
	clrSysQuickButton.clr_RibbonSystemButton_Background_Pressed = clrDefault.clr_PressedSysButton;

	clrSysQuickButton.clr_RibbonSystemButton_Border = clrDefault.clr_BorderSysButton;
	clrSysQuickButton.clr_RibbonSystemButton_Border_Disable = clrDefault.clr_BorderDisableSysButton;
	clrSysQuickButton.clr_RibbonSystemButton_Border_Highlight = clrDefault.clr_BorderHighlightSysButton;
	clrSysQuickButton.clr_RibbonSystemButton_Border_Pressed = clrDefault.clr_BorderPressedSysButton;

	clrSysQuickButton.clr_RibbonSystemButton_Text = clrDefault.clr_Text;
	clrSysQuickButton.clr_RibbonSystemButton_Text_Disable = clrDefault.clr_TextDisable;
	clrSysQuickButton.clr_RibbonSystemButton_Text_Highlight = clrDefault.clr_TextActive;
	clrSysQuickButton.clr_RibbonSystemButton_Text_Pressed = clrDefault.clr_TextPressed;

	clrSysQuickButton.clr_RibbonSystemButton_TextQuickButton = clrDefault.clr_Text;
	clrSysQuickButton.clr_RibbonSystemButton_TextQuickButton_Disable = clrDefault.clr_TextDisable;
	clrSysQuickButton.clr_RibbonSystemButton_TextQuickButton_Highlight = clrDefault.clr_TextActive;
	clrSysQuickButton.clr_RibbonSystemButton_TextQuickButton_Pressed = clrDefault.clr_TextPressed;

	/*StatusBarPane*/
	ColorTheme::ColorStatusBarPane &clrStatusBarPane = m_currentTheme.m_ColorStatusBarPane;

	clrStatusBarPane.clr_StatusBarPane_Background = clrDefault.clr_ClobalBackground;

	clrStatusBarPane.clr_StatusBarPane_Background_Highlight = RGB(255, 0, 0);
	clrStatusBarPane.clr_StatusBarPane_Background_Disable = RGB(0, 255, 0);

	clrStatusBarPane.clr_StatusBarPane_ProgressBar = RGB(255, 16, 64);
	clrStatusBarPane.clr_StatusBarPane_BarDest = RGB(8, 255, 64);

	clrStatusBarPane.clr_StatusBarPane_Text = RGB(204, 204, 204);
	clrStatusBarPane.clr_StatusBarPane_Text_Highlight = RGB(204, 204, 204);
	clrStatusBarPane.clr_StatusBarPane_Text_Disable = RGB(204, 204, 204);
	clrStatusBarPane.clr_StatusBarPane_TextProgressBar = RGB(204, 204, 204);

	/*ControlElements*/
	ColorTheme::ColorControlElements &clrCntrlElem = m_currentTheme.m_ColorContrlElem;

	clrCntrlElem.clr_RichEdit_Background = RGB(128, 128, 128);
	clrCntrlElem.clr_RichEdit_Background_Disable = RGB(128, 128, 128);
	clrCntrlElem.clr_RichEdit_Background_Highlight = RGB(128, 128, 128);

	clrCntrlElem.clr_ProgressBar_Background = clrDefault.clr_ButtonClient;
	clrCntrlElem.clr_ProgressBar_Background_Red = clrDefault.clr_ButtonClient;
	clrCntrlElem.clr_ProgressBar_Background_Orange = clrDefault.clr_ButtonClient;
	clrCntrlElem.clr_ProgressBar_Background_Yellow = clrDefault.clr_ButtonClient;
	clrCntrlElem.clr_ProgressBar_Background_Green = clrDefault.clr_ButtonClient;
	clrCntrlElem.clr_ProgressBar_Background_Blue = clrDefault.clr_ButtonClient;
	clrCntrlElem.clr_ProgressBar_Background_Indigo = clrDefault.clr_ButtonClient;
	clrCntrlElem.clr_ProgressBar_Background_Violet = clrDefault.clr_ButtonClient;

	clrCntrlElem.clr_ProgressChunk_Background = clrDefault.clr_HighlightSysButton;
	clrCntrlElem.clr_ProgressChunk_Background_Red = clrDefault.clr_HighlightSysButton;
	clrCntrlElem.clr_ProgressChunk_Background_Orange = clrDefault.clr_HighlightSysButton;
	clrCntrlElem.clr_ProgressChunk_Background_Yellow = clrDefault.clr_HighlightSysButton;
	clrCntrlElem.clr_ProgressChunk_Background_Green = clrDefault.clr_HighlightSysButton;
	clrCntrlElem.clr_ProgressChunk_Background_Blue = clrDefault.clr_HighlightSysButton;
	clrCntrlElem.clr_ProgressChunk_Background_Indigo = clrDefault.clr_HighlightSysButton;
	clrCntrlElem.clr_ProgressChunk_Background_Violet = clrDefault.clr_HighlightSysButton;

	clrCntrlElem.clr_TextHiperLink = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Red = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Orange = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Yellow = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Green = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Blue = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Indigo = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Violet = clrDefault.clr_TextHyperlink;

	clrCntrlElem.clr_TextHiperLink_Highlight = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Highlight_Red = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Highlight_Orange = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Highlight_Yellow = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Highlight_Green = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Highlight_Blue = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Highlight_Indigo = clrDefault.clr_TextHyperlink;
	clrCntrlElem.clr_TextHiperLink_Highlight_Violet = clrDefault.clr_TextHyperlink;

	/*ColorSliderZoom*/
	ColorTheme::ColorSliderZoom &clrSliderZoom = m_currentTheme.m_ColorSliderZoom;

	clrSliderZoom.clr_SliderZoom_Background = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_SliderZoom_Background_Red = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_SliderZoom_Background_Orange = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_SliderZoom_Background_Yellow = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_SliderZoom_Background_Green = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_SliderZoom_Background_Blue = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_SliderZoom_Background_Indigo = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_SliderZoom_Background_Violet = clrDefault.clr_ButtonClient;

	clrSliderZoom.clr_BottonZoom_Background = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Red = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Orange = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Yellow = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Green = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Blue = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Indigo = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Violet = clrDefault.clr_ButtonClient;

	clrSliderZoom.clr_BottonZoom_Background_Highlight = clrDefault.clr_ActiveButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Highlight_Red = clrDefault.clr_ActiveButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Highlight_Orange = clrDefault.clr_ActiveButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Highlight_Yellow = clrDefault.clr_ActiveButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Highlight_Green = clrDefault.clr_ActiveButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Highlight_Blue = clrDefault.clr_ActiveButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Highlight_Indigo = clrDefault.clr_ActiveButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Highlight_Violet = clrDefault.clr_ActiveButtonClient;

	clrSliderZoom.clr_BottonZoom_Background_Pressed = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Pressed_Red = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Pressed_Orange = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Pressed_Yellow = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Pressed_Green = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Pressed_Blue = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Pressed_Indigo = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Pressed_Violet = clrDefault.clr_PressedButtonClient;

	clrSliderZoom.clr_BottonZoom_Background_Disabled = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Disabled_Red = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Disabled_Orange = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Disabled_Yellow = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Disabled_Green = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Disabled_Blue = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Disabled_Indigo = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Disabled_Violet = clrDefault.clr_DisableButtonClient;

	clrSliderZoom.clr_BottonZoom_Background_Plus = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Red = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Orange = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Yellow = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Green = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Blue = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Indigo = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Violet = clrDefault.clr_ButtonClient;

	clrSliderZoom.clr_BottonZoom_Background_Minus = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Red = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Orange = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Yellow = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Green = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Blue = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Indigo = clrDefault.clr_ButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Violet = clrDefault.clr_ButtonClient;

	clrSliderZoom.clr_BottonZoom_Background_Plus_Disabled = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Disabled_Red = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Disabled_Orange = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Disabled_Yellow = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Disabled_Green = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Disabled_Blue = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Disabled_Indigo = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Disabled_Violet = clrDefault.clr_DisableButtonClient;

	clrSliderZoom.clr_BottonZoom_Background_Minus_Disabled = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Disabled_Red = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Disabled_Orange = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Disabled_Yellow = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Disabled_Green = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Disabled_Blue = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Disabled_Indigo = clrDefault.clr_DisableButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Disabled_Violet = clrDefault.clr_DisableButtonClient;

	clrSliderZoom.clr_BottonZoom_Background_Plus_Highlight = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Highlight_Red = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Highlight_Orange = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Highlight_Yellow = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Highlight_Green = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Highlight_Blue = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Highlight_Indigo = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Highlight_Violet = clrDefault.clr_ActiveSysButton;

	clrSliderZoom.clr_BottonZoom_Background_Minus_Highlight = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Highlight_Red = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Highlight_Orange = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Highlight_Yellow = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Highlight_Green = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Highlight_Blue = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Highlight_Indigo = clrDefault.clr_ActiveSysButton;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Highlight_Violet = clrDefault.clr_ActiveSysButton;

	clrSliderZoom.clr_BottonZoom_Background_Plus_Pressed = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Pressed_Red = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Pressed_Orange = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Pressed_Yellow = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Pressed_Green = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Pressed_Blue = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Pressed_Indigo = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Plus_Pressed_Violet = clrDefault.clr_PressedButtonClient;

	clrSliderZoom.clr_BottonZoom_Background_Minus_Pressed = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Pressed_Red = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Pressed_Orange = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Pressed_Yellow = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Pressed_Green = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Pressed_Blue = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Pressed_Indigo = clrDefault.clr_PressedButtonClient;
	clrSliderZoom.clr_BottonZoom_Background_Minus_Pressed_Violet = clrDefault.clr_PressedButtonClient;

	/*Param*/
	ColorTheme::ParametrsCMenuImage &prmDrawMenu = m_currentTheme.m_sParamCMenu;
	ColorTheme::ParametrsDrawStandard &prmDrawStandard = m_currentTheme.m_sParamStandard;
	ColorTheme::ParametrsDrawRibbon &prmDrawRibbon = m_currentTheme.m_sParamRibbon;

	m_CMIUserImage.CleanUp();

	prmDrawMenu.img_ColorImage = CMenuImages::ImageWhite;
	prmDrawMenu.clr_UserColorImageIcon = clrDefault.clr_ImageIcon;

}

void CNPUP_VisualManager::SetChildWindowStyle()
{

	CWnd *pWnd = AfxGetMainWnd();

	if (pWnd != NULL)
	{
		CWnd* pCurrent = pWnd->GetWindow(GW_CHILD);

		while (pCurrent != nullptr)
		{
			if (::IsWindow(pCurrent->GetSafeHwnd()))
			{
				if (pCurrent->IsKindOf(RUNTIME_CLASS(CTreeCtrl)))
				{
					ColorTheme::ColorThreeView clrThreeView = m_currentTheme.m_ColorThreeView;
					CTreeCtrl *pTree = (CTreeCtrl*)pCurrent;
					pTree->SetBkColor(clrThreeView.clr_ThreeViewBackground);
					pTree->SetTextColor(clrThreeView.clr_ThreeView_Text);
					pTree->SetInsertMarkColor(RGB(208, 208, 0)); // неизвестно
					pTree->SetLineColor(clrThreeView.clr_ThreeViewLine);
					::SetWindowTheme(pTree->GetSafeHwnd(), L"DarkMode_Explorer", NULL);
					pTree->Invalidate();
				}
				if (pCurrent->IsKindOf(RUNTIME_CLASS(CListBox)))
				{
					ColorTheme::ColorThreeView clrThreeView = m_currentTheme.m_ColorThreeView;
					CListBox *pList = (CListBox*)pCurrent;
				}
				if (pCurrent->IsKindOf(RUNTIME_CLASS(CScrollBar)))
				{
					CScrollBar *pScroll = (CScrollBar*)pCurrent;
					::SetWindowTheme(pScroll->GetSafeHwnd(), L"DarkMode_Explorer", NULL);
					pScroll->Invalidate();
				}
				if (pCurrent->IsKindOf(RUNTIME_CLASS(CMFCPropertyGridCtrl)))
				{
					ColorTheme::ColorPropertyGrid clrPropGrid = m_currentTheme.m_ColorPropertyGrid;
					CMFCPropertyGridCtrl *pPropGrid = (CMFCPropertyGridCtrl*)pCurrent;

					COLORREF clrBackground = clrPropGrid.clr_PropGrid_Background;
					COLORREF clrText = clrPropGrid.clr_PropGrid_Text;
					COLORREF clrGroupBackground = clrPropGrid.clr_PropGrid_GroupBackground;
					COLORREF clrGroupText = clrPropGrid.clr_PropGrid_GroupText;
					COLORREF clrDescriptionBackground = clrPropGrid.clr_PropGrid_DescriptionBackground;
					COLORREF clrDescriptionText = clrPropGrid.clr_PropGrid_DescriptionText;
					COLORREF clrLine = clrPropGrid.clr_PropGrid_Line;

					pPropGrid->SetCustomColors
					(
						clrBackground,
						clrText,
						clrGroupBackground,
						clrGroupText,
						clrDescriptionBackground,
						clrDescriptionText,
						clrLine
					);
				}
				if (pCurrent->IsKindOf(RUNTIME_CLASS(CComboBox)))
				{
					CComboBox *pCombo = (CComboBox*)pCurrent;
					::SetWindowTheme(pCombo->GetSafeHwnd(), L"DarkMode_Explorer", NULL);
					pCombo->Invalidate();
				}
				if (pCurrent->IsKindOf(RUNTIME_CLASS(CButton)))
				{
					CButton *pButton = (CButton*)pCurrent;
					::SetWindowTheme(pButton->GetSafeHwnd(), L"DarkMode_Explorer", NULL);
					pButton->Invalidate();
				}
				if (pCurrent->IsKindOf(RUNTIME_CLASS(CListBox)))
				{
					CListBox *pList = (CListBox*)pCurrent;
					::SetWindowTheme(pList->GetSafeHwnd(), L"DarkMode_Explorer", NULL);
					pList->Invalidate();
				}
				if (pCurrent->IsKindOf(RUNTIME_CLASS(CScrollView)))
				{
					CScrollView *pView = (CScrollView*)pCurrent;
					::SetWindowTheme(pView->GetSafeHwnd(), L"DarkMode_Explorer", NULL);
					pView->Invalidate();
				}

				CWnd* pNext = pCurrent->GetWindow(GW_CHILD);

				if (pNext != nullptr)
				{
					pCurrent = pNext;
					continue;
				}
			}

			while (pCurrent != nullptr && pCurrent != pWnd)
			{
				CWnd* pSibling = pCurrent->GetWindow(GW_HWNDNEXT);
				if (pSibling != nullptr)
				{
					pCurrent = pSibling;
					break;
				}
				pCurrent = pCurrent->GetParent();

				if (pCurrent == pWnd)
				{
					pCurrent = nullptr;
					break;
				}
			}
		}
	}

}

void CNPUP_VisualManager::SetStyle
(
	ThemeStyle style
)
{
	m_currentStyle = style;

	DWORD  color = 0;
	BOOL opaque = TRUE;
	COLORREF clrSystemColor;

	HRESULT hr = DwmGetColorizationColor(&color, &opaque);
	if (SUCCEEDED(hr))
	{
		clrSystemColor = RGB((color >> 16), (color >> 8), (color >> 0));
	}

	DefaultColorTheme &clrDefaultTheme = m_DefaultColor;

	clrDefaultTheme.clr_RibbonFrame_Border = RGB(128, 128, 128);
	clrDefaultTheme.clr_RibbonFrame_Border_Red = RGB(160, 64, 64);
	clrDefaultTheme.clr_RibbonFrame_Border_Orange = RGB(160, 128, 64);
	clrDefaultTheme.clr_RibbonFrame_Border_Yellow = RGB(160, 160, 64);
	clrDefaultTheme.clr_RibbonFrame_Border_Green = RGB(64, 160, 64);
	clrDefaultTheme.clr_RibbonFrame_Border_Blue = RGB(64, 64, 160);
	clrDefaultTheme.clr_RibbonFrame_Border_Indigo = RGB(128, 64, 160);
	clrDefaultTheme.clr_RibbonFrame_Border_Violet = RGB(160, 64, 128);

	switch (m_currentStyle)
	{
	case nopStyle_None:
		clrDefaultTheme.clr_ClobalBackground = RGB(16, 16, 16);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(8, 8, 8);
		clrDefaultTheme.clr_ClientBorder = RGB(16, 16, 16);

		clrDefaultTheme.clr_Caption = RGB(32, 32, 32);
		clrDefaultTheme.clr_ActiveCaption = RGB(64, 64, 64);
		clrDefaultTheme.clr_DisableCaption = RGB(8, 8, 8);

		clrDefaultTheme.clr_ButtonClient = RGB(32, 32, 32);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(64, 64, 64);
		clrDefaultTheme.clr_DisableButtonClient = RGB(8, 8, 8);
		clrDefaultTheme.clr_PressedButtonClient = RGB(8, 8, 8);

		clrDefaultTheme.clr_SysButton = RGB(16, 16, 16);
		clrDefaultTheme.clr_ActiveSysButton = RGB(64, 64, 64);
		clrDefaultTheme.clr_DisableSysButton = RGB(16, 16, 16);
		clrDefaultTheme.clr_PressedSysButton = RGB(8, 8, 8);
		clrDefaultTheme.clr_HighlightSysButton = RGB(128, 128, 128);

		clrDefaultTheme.clr_ImageIcon = RGB(255, 255, 255);

		clrDefaultTheme.clr_Text = RGB(224, 224, 244);
		clrDefaultTheme.clr_TextDisable = RGB(128, 128, 128);
		clrDefaultTheme.clr_TextActive = RGB(255, 255, 255);
		clrDefaultTheme.clr_TextPressed = RGB(128, 128, 128);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);

		break;
	case nopStyle_DarkGray:
		clrDefaultTheme.clr_ClobalBackground = RGB(24, 24, 24);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(16, 16, 16);
		clrDefaultTheme.clr_ClientBorder = RGB(24, 24, 24);

		clrDefaultTheme.clr_Caption = RGB(32, 32, 32);
		clrDefaultTheme.clr_ActiveCaption = RGB(64, 64, 64);
		clrDefaultTheme.clr_DisableCaption = RGB(16, 16, 16);

		clrDefaultTheme.clr_ButtonClient = RGB(16, 16, 16);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(128, 128, 128);
		clrDefaultTheme.clr_DisableButtonClient = RGB(16, 16, 16);
		clrDefaultTheme.clr_PressedButtonClient = RGB(8, 8, 8);

		clrDefaultTheme.clr_SysButton = RGB(24, 24, 24);
		clrDefaultTheme.clr_ActiveSysButton = RGB(64, 64, 64);
		clrDefaultTheme.clr_DisableSysButton = RGB(24, 24, 24);
		clrDefaultTheme.clr_PressedSysButton = RGB(8, 8, 8);
		clrDefaultTheme.clr_HighlightSysButton = RGB(128, 128, 128);

		clrDefaultTheme.clr_BorderButtonClient = RGB(16, 16, 16);
		clrDefaultTheme.clr_BorderActiveButtonClient = RGB(128, 128, 128);
		clrDefaultTheme.clr_BorderDisableButtonClient = RGB(16, 16, 16);
		clrDefaultTheme.clr_BorderPressedButtonClient = RGB(8, 8, 8);

		clrDefaultTheme.clr_BorderSysButton = RGB(64, 64, 64);
		clrDefaultTheme.clr_BorderActiveSysButton = RGB(128, 128, 128);
		clrDefaultTheme.clr_BorderDisableSysButton = RGB(64, 64, 64);
		clrDefaultTheme.clr_BorderPressedSysButton = RGB(128, 128, 128);
		clrDefaultTheme.clr_BorderHighlightSysButton = RGB(236, 236, 236);

		clrDefaultTheme.clr_RibbonFrame = RGB(80, 80, 80);
		clrDefaultTheme.clr_RibbonFrame_Red = RGB(96, 0, 0);
		clrDefaultTheme.clr_RibbonFrame_Orange = RGB(96, 64, 0);
		clrDefaultTheme.clr_RibbonFrame_Yellow = RGB(96, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Green = RGB(0, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Blue = RGB(0, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Indigo = RGB(64, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Violet = RGB(96, 0, 64);

		clrDefaultTheme.clr_ImageIcon = RGB(204, 204, 204);

		clrDefaultTheme.clr_Text = RGB(224, 224, 224);
		clrDefaultTheme.clr_TextDisable = RGB(128, 128, 128);
		clrDefaultTheme.clr_TextActive = RGB(255, 255, 255);
		clrDefaultTheme.clr_TextPressed = RGB(128, 128, 128);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);


		break;
	case nopStyle_White:
		clrDefaultTheme.clr_ClobalBackground = RGB(236, 236, 236);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(180, 180, 180);
		clrDefaultTheme.clr_ClientBorder = RGB(196, 196, 196);

		clrDefaultTheme.clr_Caption = RGB(204, 204, 204);
		clrDefaultTheme.clr_ActiveCaption = RGB(255, 255, 255);
		clrDefaultTheme.clr_DisableCaption = RGB(180, 180, 180);

		clrDefaultTheme.clr_ButtonClient = RGB(180, 180, 180);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(196, 196, 196);
		clrDefaultTheme.clr_DisableButtonClient = RGB(180, 180, 180);
		clrDefaultTheme.clr_PressedButtonClient = RGB(160, 160, 160);

		clrDefaultTheme.clr_SysButton = RGB(236, 236, 236);
		clrDefaultTheme.clr_ActiveSysButton = RGB(196, 196, 196);
		clrDefaultTheme.clr_DisableSysButton = RGB(236, 236, 236);
		clrDefaultTheme.clr_PressedSysButton = RGB(128, 128, 128);
		clrDefaultTheme.clr_HighlightSysButton = RGB(196, 196, 196);

		clrDefaultTheme.clr_BorderButtonClient = RGB(180, 180, 180);
		clrDefaultTheme.clr_BorderActiveButtonClient = RGB(196, 196, 196);
		clrDefaultTheme.clr_BorderDisableButtonClient = RGB(180, 180, 180);
		clrDefaultTheme.clr_BorderPressedButtonClient = RGB(160, 160, 160);

		clrDefaultTheme.clr_BorderSysButton = RGB(32, 32, 32);
		clrDefaultTheme.clr_BorderActiveSysButton = RGB(80, 80, 80);
		clrDefaultTheme.clr_BorderDisableSysButton = RGB(32, 32, 32);
		clrDefaultTheme.clr_BorderPressedSysButton = RGB(0, 0, 0);
		clrDefaultTheme.clr_BorderHighlightSysButton = RGB(128, 128, 128);

		clrDefaultTheme.clr_RibbonFrame = RGB(128, 128, 128);
		clrDefaultTheme.clr_RibbonFrame_Red = RGB(160, 64, 64);
		clrDefaultTheme.clr_RibbonFrame_Orange = RGB(160, 128, 64);
		clrDefaultTheme.clr_RibbonFrame_Yellow = RGB(160, 160, 64);
		clrDefaultTheme.clr_RibbonFrame_Green = RGB(64, 160, 64);
		clrDefaultTheme.clr_RibbonFrame_Blue = RGB(64, 64, 160);
		clrDefaultTheme.clr_RibbonFrame_Indigo = RGB(128, 64, 160);
		clrDefaultTheme.clr_RibbonFrame_Violet = RGB(160, 64, 128);


		clrDefaultTheme.clr_ImageIcon = RGB(0, 0, 0);

		clrDefaultTheme.clr_Text = RGB(0, 0, 0);
		clrDefaultTheme.clr_TextDisable = RGB(64, 64, 64);
		clrDefaultTheme.clr_TextActive = RGB(0, 0, 0);
		clrDefaultTheme.clr_TextPressed = RGB(64, 64, 64);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);

		break;
	case nopStyle_Blue:
		clrDefaultTheme.clr_ClobalBackground = RGB(8, 32, 64);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(0, 8, 16);
		clrDefaultTheme.clr_ClientBorder = RGB(8, 16, 32);

		clrDefaultTheme.clr_Caption = RGB(8, 32, 64);
		clrDefaultTheme.clr_ActiveCaption = RGB(16, 64, 128);
		clrDefaultTheme.clr_DisableCaption = RGB(0, 8, 16);

		clrDefaultTheme.clr_ButtonClient = RGB(0, 8, 16);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(8, 32, 64);
		clrDefaultTheme.clr_DisableButtonClient = RGB(0, 8, 16);
		clrDefaultTheme.clr_PressedButtonClient = RGB(0, 8, 16);

		clrDefaultTheme.clr_SysButton = RGB(8, 32, 64);
		clrDefaultTheme.clr_ActiveSysButton = RGB(8, 48, 96);
		clrDefaultTheme.clr_DisableSysButton = RGB(8, 32, 64);
		clrDefaultTheme.clr_PressedSysButton = RGB(0, 8, 16);
		clrDefaultTheme.clr_HighlightSysButton = RGB(0, 128, 204);

		clrDefaultTheme.clr_BorderButtonClient = RGB(0, 8, 16);
		clrDefaultTheme.clr_BorderActiveButtonClient = RGB(8, 32, 64);
		clrDefaultTheme.clr_BorderDisableButtonClient = RGB(0, 8, 16);
		clrDefaultTheme.clr_BorderPressedButtonClient = RGB(0, 8, 16);

		clrDefaultTheme.clr_BorderSysButton = RGB(8, 32, 128);
		clrDefaultTheme.clr_BorderActiveSysButton = RGB(8, 48, 128);
		clrDefaultTheme.clr_BorderDisableSysButton = RGB(8, 32, 64);
		clrDefaultTheme.clr_BorderPressedSysButton = RGB(0, 8, 16);
		clrDefaultTheme.clr_BorderHighlightSysButton = RGB(0, 128, 255);

		clrDefaultTheme.clr_RibbonFrame = RGB(16, 64, 128);
		clrDefaultTheme.clr_RibbonFrame_Red = RGB(96, 0, 0);
		clrDefaultTheme.clr_RibbonFrame_Orange = RGB(96, 64, 0);
		clrDefaultTheme.clr_RibbonFrame_Yellow = RGB(96, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Green = RGB(0, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Blue = RGB(0, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Indigo = RGB(64, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Violet = RGB(96, 0, 64);

		clrDefaultTheme.clr_ImageIcon = RGB(224, 224, 224);

		clrDefaultTheme.clr_Text = RGB(224, 224, 224);
		clrDefaultTheme.clr_TextDisable = RGB(32, 64, 128);
		clrDefaultTheme.clr_TextActive = RGB(255, 255, 255);
		clrDefaultTheme.clr_TextPressed = RGB(32, 64, 128);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);

		break;
	case nopStyle_Green:
		clrDefaultTheme.clr_ClobalBackground = RGB(0, 32, 8);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(0, 16, 4);
		clrDefaultTheme.clr_ClientBorder = RGB(0, 32, 8);

		clrDefaultTheme.clr_Caption = RGB(0, 32, 8);
		clrDefaultTheme.clr_ActiveCaption = RGB(0, 128, 64);
		clrDefaultTheme.clr_DisableCaption = RGB(0, 8, 0);

		clrDefaultTheme.clr_ButtonClient = RGB(0, 16, 4);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(0, 64, 16);
		clrDefaultTheme.clr_DisableButtonClient = RGB(0, 16, 4);
		clrDefaultTheme.clr_PressedButtonClient = RGB(0, 16, 4);

		clrDefaultTheme.clr_SysButton = RGB(0, 32, 8);
		clrDefaultTheme.clr_ActiveSysButton = RGB(0, 64, 16);
		clrDefaultTheme.clr_DisableSysButton = RGB(0, 32, 8);
		clrDefaultTheme.clr_PressedSysButton = RGB(0, 16, 4);

		clrDefaultTheme.clr_RibbonFrame = RGB(0, 128, 64);
		clrDefaultTheme.clr_RibbonFrame_Red = RGB(96, 0, 0);
		clrDefaultTheme.clr_RibbonFrame_Orange = RGB(96, 64, 0);
		clrDefaultTheme.clr_RibbonFrame_Yellow = RGB(96, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Green = RGB(0, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Blue = RGB(0, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Indigo = RGB(64, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Violet = RGB(96, 0, 64);

		clrDefaultTheme.clr_ImageIcon = RGB(224, 224, 224);

		clrDefaultTheme.clr_Text = RGB(224, 224, 224);
		clrDefaultTheme.clr_TextDisable = RGB(0, 128, 0);
		clrDefaultTheme.clr_TextActive = RGB(255, 255, 255);
		clrDefaultTheme.clr_TextPressed = RGB(0, 128, 0);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);

		break;
	case nopStyle_Purple:
		clrDefaultTheme.clr_ClobalBackground = RGB(32, 8, 32);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(16, 0, 16);
		clrDefaultTheme.clr_ClientBorder = RGB(8, 0, 8);

		clrDefaultTheme.clr_Caption = RGB(32, 8, 32);
		clrDefaultTheme.clr_ActiveCaption = RGB(80, 0, 80);
		clrDefaultTheme.clr_DisableCaption = RGB(8, 0, 8);

		clrDefaultTheme.clr_ButtonClient = RGB(16, 0, 16);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(80, 16, 80);
		clrDefaultTheme.clr_DisableButtonClient = RGB(16, 0, 16);
		clrDefaultTheme.clr_PressedButtonClient = RGB(32, 8, 32);

		clrDefaultTheme.clr_SysButton = RGB(32, 8, 32);
		clrDefaultTheme.clr_ActiveSysButton = RGB(64, 16, 64);
		clrDefaultTheme.clr_DisableSysButton = RGB(32, 8, 32);
		clrDefaultTheme.clr_PressedSysButton = RGB(16, 0, 16);
		clrDefaultTheme.clr_HighlightSysButton = RGB(128, 32, 128);

		clrDefaultTheme.clr_RibbonFrame = RGB(80, 0, 80);
		clrDefaultTheme.clr_RibbonFrame_Red = RGB(96, 0, 0);
		clrDefaultTheme.clr_RibbonFrame_Orange = RGB(96, 64, 0);
		clrDefaultTheme.clr_RibbonFrame_Yellow = RGB(96, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Green = RGB(0, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Blue = RGB(0, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Indigo = RGB(64, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Violet = RGB(96, 0, 64);

		clrDefaultTheme.clr_ImageIcon = RGB(224, 224, 224);

		clrDefaultTheme.clr_Text = RGB(224, 224, 224);
		clrDefaultTheme.clr_TextDisable = RGB(128, 0, 128);
		clrDefaultTheme.clr_TextActive = RGB(255, 255, 255);
		clrDefaultTheme.clr_TextPressed = RGB(128, 0, 128);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);

		break;
	case nopStyle_DarkBlue:
		clrDefaultTheme.clr_ClobalBackground = RGB(16, 16, 16);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(8, 8, 8);
		clrDefaultTheme.clr_ClientBorder = RGB(16, 16, 16);

		clrDefaultTheme.clr_Caption = RGB(32, 32, 32);
		clrDefaultTheme.clr_ActiveCaption = RGB(16, 64, 128);
		clrDefaultTheme.clr_DisableCaption = RGB(8, 8, 8);

		clrDefaultTheme.clr_ButtonClient = RGB(8, 8, 8);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(8, 32, 64);
		clrDefaultTheme.clr_DisableButtonClient = RGB(8, 8, 8);
		clrDefaultTheme.clr_PressedButtonClient = RGB(0, 8, 16);

		clrDefaultTheme.clr_SysButton = RGB(16, 16, 16);
		clrDefaultTheme.clr_ActiveSysButton = RGB(8, 48, 96);
		clrDefaultTheme.clr_DisableSysButton = RGB(16, 16, 16);
		clrDefaultTheme.clr_PressedSysButton = RGB(0, 8, 16);
		clrDefaultTheme.clr_HighlightSysButton = RGB(0, 128, 204);

		clrDefaultTheme.clr_RibbonFrame = RGB(16, 64, 128);
		clrDefaultTheme.clr_RibbonFrame_Red = RGB(96, 0, 0);
		clrDefaultTheme.clr_RibbonFrame_Orange = RGB(96, 64, 0);
		clrDefaultTheme.clr_RibbonFrame_Yellow = RGB(96, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Green = RGB(0, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Blue = RGB(0, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Indigo = RGB(64, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Violet = RGB(96, 0, 64);

		clrDefaultTheme.clr_ImageIcon = RGB(128, 204, 255);

		clrDefaultTheme.clr_Text = RGB(224, 224, 224);
		clrDefaultTheme.clr_TextDisable = RGB(128, 128, 128);
		clrDefaultTheme.clr_TextActive = RGB(255, 255, 255);
		clrDefaultTheme.clr_TextPressed = RGB(128, 128, 128);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);

		break;
	case nopStyle_DarkGreen:
		clrDefaultTheme.clr_ClobalBackground = RGB(16, 16, 16);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(8, 8, 8);
		clrDefaultTheme.clr_ClientBorder = RGB(16, 16, 16);

		clrDefaultTheme.clr_Caption = RGB(32, 32, 32);
		clrDefaultTheme.clr_ActiveCaption = RGB(0, 64, 0);
		clrDefaultTheme.clr_DisableCaption = RGB(8, 8, 8);

		clrDefaultTheme.clr_ButtonClient = RGB(8, 8, 8);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(0, 128, 0);
		clrDefaultTheme.clr_DisableButtonClient = RGB(8, 8, 8);
		clrDefaultTheme.clr_PressedButtonClient = RGB(0, 32, 0);

		clrDefaultTheme.clr_SysButton = RGB(16, 16, 16);
		clrDefaultTheme.clr_ActiveSysButton = RGB(0, 64, 0);
		clrDefaultTheme.clr_DisableSysButton = RGB(16, 16, 16);
		clrDefaultTheme.clr_PressedSysButton = RGB(8, 8, 8);
		clrDefaultTheme.clr_HighlightSysButton = RGB(0, 128, 0);

		clrDefaultTheme.clr_RibbonFrame = RGB(0, 64, 0);
		clrDefaultTheme.clr_RibbonFrame_Red = RGB(96, 0, 0);
		clrDefaultTheme.clr_RibbonFrame_Orange = RGB(96, 64, 0);
		clrDefaultTheme.clr_RibbonFrame_Yellow = RGB(96, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Green = RGB(0, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Blue = RGB(0, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Indigo = RGB(64, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Violet = RGB(96, 0, 64);

		clrDefaultTheme.clr_ImageIcon = RGB(0, 255, 0);

		clrDefaultTheme.clr_Text = RGB(224, 224, 224);
		clrDefaultTheme.clr_TextDisable = RGB(128, 128, 128);
		clrDefaultTheme.clr_TextActive = RGB(255, 255, 255);
		clrDefaultTheme.clr_TextPressed = RGB(128, 128, 128);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);

		break;
	case nopStyle_DarkYollow:
		clrDefaultTheme.clr_ClobalBackground = RGB(16, 16, 16);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(8, 8, 8);
		clrDefaultTheme.clr_ClientBorder = RGB(16, 16, 16);

		clrDefaultTheme.clr_Caption = RGB(32, 32, 32);
		clrDefaultTheme.clr_ActiveCaption = RGB(128, 128, 0);
		clrDefaultTheme.clr_DisableCaption = RGB(8, 8, 8);

		clrDefaultTheme.clr_ButtonClient = RGB(8, 8, 8);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(128, 128, 0);
		clrDefaultTheme.clr_DisableButtonClient = RGB(8, 8, 0);
		clrDefaultTheme.clr_PressedButtonClient = RGB(32, 32, 8);

		clrDefaultTheme.clr_SysButton = RGB(16, 16, 16);
		clrDefaultTheme.clr_ActiveSysButton = RGB(128, 128, 0);
		clrDefaultTheme.clr_DisableSysButton = RGB(16, 16, 16);
		clrDefaultTheme.clr_PressedSysButton = RGB(8, 8, 0);
		clrDefaultTheme.clr_HighlightSysButton = RGB(204, 204, 0);

		clrDefaultTheme.clr_RibbonFrame = RGB(128, 128, 0);
		clrDefaultTheme.clr_RibbonFrame_Red = RGB(96, 0, 0);
		clrDefaultTheme.clr_RibbonFrame_Orange = RGB(96, 64, 0);
		clrDefaultTheme.clr_RibbonFrame_Yellow = RGB(96, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Green = RGB(0, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Blue = RGB(0, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Indigo = RGB(64, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Violet = RGB(96, 0, 64);

		clrDefaultTheme.clr_ImageIcon = RGB(255, 255, 0);

		clrDefaultTheme.clr_Text = RGB(224, 224, 224);
		clrDefaultTheme.clr_TextDisable = RGB(128, 128, 128);
		clrDefaultTheme.clr_TextActive = RGB(255, 255, 255);
		clrDefaultTheme.clr_TextPressed = RGB(128, 128, 128);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);

		break;
	case nopStyle_DarkPurple:
		clrDefaultTheme.clr_ClobalBackground = RGB(16, 16, 16);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(8, 8, 8);
		clrDefaultTheme.clr_ClientBorder = RGB(16, 16, 16);

		clrDefaultTheme.clr_Caption = RGB(32, 32, 32);
		clrDefaultTheme.clr_ActiveCaption = RGB(80, 0, 80);
		clrDefaultTheme.clr_DisableCaption = RGB(8, 8, 8);

		clrDefaultTheme.clr_ButtonClient = RGB(8, 8, 8);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(80, 16, 80);
		clrDefaultTheme.clr_DisableButtonClient = RGB(8, 8, 8);
		clrDefaultTheme.clr_PressedButtonClient = RGB(32, 8, 32);

		clrDefaultTheme.clr_SysButton = RGB(16, 16, 16);
		clrDefaultTheme.clr_ActiveSysButton = RGB(64, 16, 64);
		clrDefaultTheme.clr_DisableSysButton = RGB(16, 16, 16);
		clrDefaultTheme.clr_PressedSysButton = RGB(8, 8, 8);
		clrDefaultTheme.clr_HighlightSysButton = RGB(128, 128, 128);

		clrDefaultTheme.clr_RibbonFrame = RGB(80, 0, 80);
		clrDefaultTheme.clr_RibbonFrame_Red = RGB(96, 0, 0);
		clrDefaultTheme.clr_RibbonFrame_Orange = RGB(96, 64, 0);
		clrDefaultTheme.clr_RibbonFrame_Yellow = RGB(96, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Green = RGB(0, 96, 0);
		clrDefaultTheme.clr_RibbonFrame_Blue = RGB(0, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Indigo = RGB(64, 0, 96);
		clrDefaultTheme.clr_RibbonFrame_Violet = RGB(96, 0, 64);

		clrDefaultTheme.clr_ImageIcon = RGB(244, 0, 244);

		clrDefaultTheme.clr_Text = RGB(224, 224, 224);
		clrDefaultTheme.clr_TextDisable = RGB(128, 128, 128);
		clrDefaultTheme.clr_TextActive = RGB(255, 255, 255);
		clrDefaultTheme.clr_TextPressed = RGB(128, 128, 128);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);

		break;
	case nopStyle_WhiteBlue:
		clrDefaultTheme.clr_ClobalBackground = RGB(180, 180, 180);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(236, 236, 236);
		clrDefaultTheme.clr_ClientBorder = RGB(128, 128, 128);

		clrDefaultTheme.clr_Caption = RGB(128, 128, 128);
		clrDefaultTheme.clr_ActiveCaption = RGB(64, 128, 196);
		clrDefaultTheme.clr_DisableCaption = RGB(128, 128, 128);

		clrDefaultTheme.clr_ButtonClient = RGB(236, 236, 236);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(128, 196, 255);
		clrDefaultTheme.clr_DisableButtonClient = RGB(236, 236, 236);
		clrDefaultTheme.clr_PressedButtonClient = RGB(32, 64, 160);

		clrDefaultTheme.clr_SysButton = RGB(180, 180, 180);
		clrDefaultTheme.clr_ActiveSysButton = RGB(64, 128, 196);
		clrDefaultTheme.clr_DisableSysButton = RGB(180, 180, 180);
		clrDefaultTheme.clr_PressedSysButton = RGB(32, 64, 128);
		clrDefaultTheme.clr_HighlightSysButton = RGB(128, 196, 255);

		clrDefaultTheme.clr_RibbonFrame = RGB(64, 128, 196);
		clrDefaultTheme.clr_RibbonFrame_Red = RGB(160, 64, 64);
		clrDefaultTheme.clr_RibbonFrame_Orange = RGB(160, 128, 64);
		clrDefaultTheme.clr_RibbonFrame_Yellow = RGB(160, 160, 64);
		clrDefaultTheme.clr_RibbonFrame_Green = RGB(64, 160, 64);
		clrDefaultTheme.clr_RibbonFrame_Blue = RGB(64, 64, 160);
		clrDefaultTheme.clr_RibbonFrame_Indigo = RGB(128, 64, 160);
		clrDefaultTheme.clr_RibbonFrame_Violet = RGB(160, 64, 128);

		clrDefaultTheme.clr_ImageIcon = RGB(0, 32, 128);

		clrDefaultTheme.clr_Text = RGB(0, 0, 0);
		clrDefaultTheme.clr_TextDisable = RGB(64, 64, 64);
		clrDefaultTheme.clr_TextActive = RGB(0, 0, 0);
		clrDefaultTheme.clr_TextPressed = RGB(64, 64, 64);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);

		break;
	case nopStyle_WhiteGreen:
		clrDefaultTheme.clr_ClobalBackground = RGB(180, 180, 180);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(236, 236, 236);
		clrDefaultTheme.clr_ClientBorder = RGB(128, 128, 128);

		clrDefaultTheme.clr_Caption = RGB(128, 128, 128);
		clrDefaultTheme.clr_ActiveCaption = RGB(64, 196, 128);
		clrDefaultTheme.clr_DisableCaption = RGB(128, 128, 128);

		clrDefaultTheme.clr_ButtonClient = RGB(236, 236, 236);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(128, 255, 196);
		clrDefaultTheme.clr_DisableButtonClient = RGB(236, 236, 236);
		clrDefaultTheme.clr_PressedButtonClient = RGB(32, 64, 160);

		clrDefaultTheme.clr_SysButton = RGB(180, 180, 180);
		clrDefaultTheme.clr_ActiveSysButton = RGB(64, 196, 128);
		clrDefaultTheme.clr_DisableSysButton = RGB(180, 180, 180);
		clrDefaultTheme.clr_PressedSysButton = RGB(32, 128, 64);
		clrDefaultTheme.clr_HighlightSysButton = RGB(128, 255, 196);

		clrDefaultTheme.clr_RibbonFrame = RGB(64, 196, 128);
		clrDefaultTheme.clr_RibbonFrame_Red = RGB(160, 64, 64);
		clrDefaultTheme.clr_RibbonFrame_Orange = RGB(160, 128, 64);
		clrDefaultTheme.clr_RibbonFrame_Yellow = RGB(160, 160, 64);
		clrDefaultTheme.clr_RibbonFrame_Green = RGB(64, 160, 64);
		clrDefaultTheme.clr_RibbonFrame_Blue = RGB(64, 64, 160);
		clrDefaultTheme.clr_RibbonFrame_Indigo = RGB(128, 64, 160);
		clrDefaultTheme.clr_RibbonFrame_Violet = RGB(160, 64, 128);

		clrDefaultTheme.clr_ImageIcon = RGB(0, 128, 32);

		clrDefaultTheme.clr_Text = RGB(0, 0, 0);
		clrDefaultTheme.clr_TextDisable = RGB(64, 64, 64);
		clrDefaultTheme.clr_TextActive = RGB(0, 0, 0);
		clrDefaultTheme.clr_TextPressed = RGB(64, 64, 64);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);

		break;
	case nopStyle_WhiteYollow:
		clrDefaultTheme.clr_ClobalBackground = RGB(180, 180, 180);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(236, 236, 236);
		clrDefaultTheme.clr_ClientBorder = RGB(128, 128, 128);

		clrDefaultTheme.clr_Caption = RGB(128, 128, 128);
		clrDefaultTheme.clr_ActiveCaption = RGB(196, 196, 64);
		clrDefaultTheme.clr_DisableCaption = RGB(128, 128, 128);

		clrDefaultTheme.clr_ButtonClient = RGB(236, 236, 236);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(255, 255, 128);
		clrDefaultTheme.clr_DisableButtonClient = RGB(236, 236, 236);
		clrDefaultTheme.clr_PressedButtonClient = RGB(160, 160, 16);

		clrDefaultTheme.clr_SysButton = RGB(180, 180, 180);
		clrDefaultTheme.clr_ActiveSysButton = RGB(196, 196, 64);
		clrDefaultTheme.clr_DisableSysButton = RGB(180, 180, 180);
		clrDefaultTheme.clr_PressedSysButton = RGB(128, 128, 32);
		clrDefaultTheme.clr_HighlightSysButton = RGB(255, 255, 128);

		clrDefaultTheme.clr_RibbonFrame = RGB(196, 196, 64);
		clrDefaultTheme.clr_RibbonFrame_Red = RGB(160, 64, 64);
		clrDefaultTheme.clr_RibbonFrame_Orange = RGB(160, 128, 64);
		clrDefaultTheme.clr_RibbonFrame_Yellow = RGB(160, 160, 64);
		clrDefaultTheme.clr_RibbonFrame_Green = RGB(64, 160, 64);
		clrDefaultTheme.clr_RibbonFrame_Blue = RGB(64, 64, 160);
		clrDefaultTheme.clr_RibbonFrame_Indigo = RGB(128, 64, 160);
		clrDefaultTheme.clr_RibbonFrame_Violet = RGB(160, 64, 128);

		clrDefaultTheme.clr_ImageIcon = RGB(204, 204, 0);

		clrDefaultTheme.clr_Text = RGB(0, 0, 0);
		clrDefaultTheme.clr_TextDisable = RGB(64, 64, 64);
		clrDefaultTheme.clr_TextActive = RGB(0, 0, 0);
		clrDefaultTheme.clr_TextPressed = RGB(64, 64, 64);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);

		break;
	case nopStyle_WhitePurple:
		clrDefaultTheme.clr_ClobalBackground = RGB(180, 180, 180);
		clrDefaultTheme.clr_ClientFrameBackground = RGB(236, 236, 236);
		clrDefaultTheme.clr_ClientBorder = RGB(128, 128, 128);

		clrDefaultTheme.clr_Caption = RGB(128, 128, 128);
		clrDefaultTheme.clr_ActiveCaption = RGB(196, 64, 196);
		clrDefaultTheme.clr_DisableCaption = RGB(128, 128, 128);

		clrDefaultTheme.clr_ButtonClient = RGB(236, 236, 236);
		clrDefaultTheme.clr_ActiveButtonClient = RGB(255, 128, 255);
		clrDefaultTheme.clr_DisableButtonClient = RGB(236, 236, 236);
		clrDefaultTheme.clr_PressedButtonClient = RGB(160, 160, 16);

		clrDefaultTheme.clr_SysButton = RGB(180, 180, 180);
		clrDefaultTheme.clr_ActiveSysButton = RGB(128, 32, 128);
		clrDefaultTheme.clr_DisableSysButton = RGB(180, 180, 180);
		clrDefaultTheme.clr_PressedSysButton = RGB(80, 32, 80);
		clrDefaultTheme.clr_HighlightSysButton = RGB(255, 32, 255);

		clrDefaultTheme.clr_RibbonFrame = RGB(196, 64, 196);
		clrDefaultTheme.clr_RibbonFrame_Red = RGB(160, 64, 64);
		clrDefaultTheme.clr_RibbonFrame_Orange = RGB(160, 128, 64);
		clrDefaultTheme.clr_RibbonFrame_Yellow = RGB(160, 160, 64);
		clrDefaultTheme.clr_RibbonFrame_Green = RGB(64, 160, 64);
		clrDefaultTheme.clr_RibbonFrame_Blue = RGB(64, 64, 160);
		clrDefaultTheme.clr_RibbonFrame_Indigo = RGB(128, 64, 160);
		clrDefaultTheme.clr_RibbonFrame_Violet = RGB(160, 64, 128);

		clrDefaultTheme.clr_ImageIcon = RGB(204, 0, 204);

		clrDefaultTheme.clr_Text = RGB(0, 0, 0);
		clrDefaultTheme.clr_TextDisable = RGB(64, 64, 64);
		clrDefaultTheme.clr_TextActive = RGB(0, 0, 0);
		clrDefaultTheme.clr_TextPressed = RGB(64, 64, 64);
		clrDefaultTheme.clr_TextHyperlink = RGB(128, 128, 255);

		break;
	}

	SetColorThemeGlobal();
	SetChildWindowStyle();

	RedrawAll();

}

void CNPUP_VisualManager::OnUpdateSystemColors()
{
	/*заглушка*/
}

BOOL CNPUP_VisualManager::OnNcPaint(CWnd * pWnd, const CObList & lstSysButtons, CRect rectRedraw)
{
	return FALSE;
}

BOOL CNPUP_VisualManager::OnNcActivate(CWnd * pWnd, BOOL bActive)
{
	return FALSE;
}

#pragma endregion



// ============================================================================
// КЛАССИЧЕСКИЕ ЭЛЕМЕНТЫ ИНТЕРФЕЙСА / STANDARD UI FUNCTIONS
// ----------------------------------------------------------------------------
// Переопределение отрисовки стандартных объектов Win32 и MFC (Classic UI).
// Включает меню, классические панели инструментов (Toolbars) и строки 
// состояния.
//
// Overrides for standard Win32 and MFC interface elements (Classic UI).
// Covers menus, classic toolbars, and status bars.
// ============================================================================
#pragma region Standard

#pragma region MainFrame
void CNPUP_VisualManager::OnFillBarBackground
(
	CDC * pDC,
	CBasePane * pBar,
	CRect rectClient,
	CRect rectClip,
	BOOL bNCArea
)
{
	ColorTheme::ColorToolbars toolBarClr = m_currentTheme.m_ColorToolbars;
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	ColorTheme::ColorTabs clrTabs = m_currentTheme.m_ColorTabs;
	ColorTheme::ColorStatusBarPane clrStatusBar = m_currentTheme.m_ColorStatusBarPane;
	ColorTheme::ColorPaneCaptionsAndButton clrCaptionBar = m_currentTheme.m_ColorPaneCaptionsAndButton;
	ColorTheme::ColorRibbonBackground clrRibbonBar = m_currentTheme.m_ColorRibbonBackground;

	COLORREF clrBorder = toolBarClr.clr_BorderToolbar; //рамка
	COLORREF clrBackground = toolBarClr.clr_BackgroundToolbar; //фон

	if (pBar != nullptr && pBar->IsKindOf(RUNTIME_CLASS(CMFCPopupMenuBar)))
	{
		clrBackground = RGB(255, 0, 0);
	}

	if (pBar != nullptr && pBar->IsKindOf(RUNTIME_CLASS(CMFCRibbonPanelMenuBar)))
	{
		clrBackground = RGB(255, 0, 0);
	}

	if (pBar != nullptr && pBar->IsKindOf(RUNTIME_CLASS(CMFCRibbonBar)))
	{
		clrBackground = clrRibbonBar.clr_RibbonBackground;
	}

	if (pBar != nullptr && pBar->IsKindOf(RUNTIME_CLASS(CMFCCaptionBar)))
	{
		clrBackground = clrCaptionBar.clr_PaneCaption_BackgroundFrame;
	}

	if (pBar != nullptr && pBar->IsKindOf(RUNTIME_CLASS(CMFCRibbonStatusBar)))
	{
		clrBackground = clrStatusBar.clr_StatusBarPane_Background;
	}

	if (pBar != nullptr && pBar->IsKindOf(RUNTIME_CLASS(CMFCAutoHideBar)))
	{
		clrBackground = clrTabs.clr_BackgroundFrameDockHide;
	}

	if (pBar != nullptr && pBar->IsKindOf(RUNTIME_CLASS(CAutoHideDockSite)))
	{
		clrBackground = clrTabs.clr_BackgroundFrameDockHide;
	}

	if (pBar != nullptr && pBar->IsKindOf(RUNTIME_CLASS(CMFCToolBar)))
	{
		if (pBar != nullptr && pBar->IsKindOf(RUNTIME_CLASS(CMFCPopupMenuBar)))
		{
			clrBackground = clrGlblElem.clr_BackgroundMenu;
		}
		else
		{
			clrBackground = toolBarClr.clr_BackgroundToolbar;
		}
	}

	if (pBar != nullptr && pBar->IsKindOf(RUNTIME_CLASS(CMFCColorBar)))
	{
		clrBackground = RGB(255, 0, 255);
	}


	pDC->FillSolidRect(rectClient, clrBorder);
	pDC->FillSolidRect(rectClip, clrBackground);
}

COLORREF CNPUP_VisualManager::GetToolbarButtonTextColor
(
	CMFCToolBarButton * pButton,
	CMFCVisualManager::AFX_BUTTON_STATE state
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	COLORREF clrTextColor = clrGlblElem.clr_TextMenu;

	if (state == ButtonsIsRegular)
	{
		clrTextColor = clrGlblElem.clr_TextMenu;
	}
	else if (state == ButtonsIsPressed)
	{
		clrTextColor = clrGlblElem.clr_TextMenu_Pressed;
	}
	else if (state == ButtonsIsHighlighted)
	{
		clrTextColor = clrGlblElem.clr_TextMenu_Highlight;
	}

	return clrTextColor;
}

COLORREF CNPUP_VisualManager::GetMenuItemTextColor
(
	CMFCToolBarMenuButton * pButton,
	BOOL bHighlighted,
	BOOL bDisabled
)
{

	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;

	COLORREF clrTextColor = RGB(0, 0, 0);

	if (bHighlighted)
	{
		if (bDisabled)
		{
			clrTextColor = clrGlblElem.clr_TextMenu_Disable;
		}
		else
		{
			clrTextColor = clrGlblElem.clr_TextMenu_Highlight;
		}
	}
	else
	{
		if (bDisabled)
		{
			clrTextColor = clrGlblElem.clr_TextMenu_Disable;
		}
		else
		{
			clrTextColor = clrGlblElem.clr_TextMenu;
		}
	}

	return clrTextColor;
}

COLORREF CNPUP_VisualManager::OnDrawPaneCaption
(
	CDC * pDC,
	CDockablePane * pBar,
	BOOL bActive,
	CRect rectCaption,
	CRect rectButtons
)
{
	ColorTheme::ColorMiniFrame minFrameClr = m_currentTheme.m_ColorMiniFrame;

	COLORREF clrTextPaneCaption = RGB(0, 255, 0);
	COLORREF clrButtonPaneCaption = RGB(0, 0, 255);
	COLORREF clrBackgroundPaneCaption = RGB(255, 0, 0);

	if (bActive)
	{
		clrButtonPaneCaption = minFrameClr.clr_BackgroundMiniFrame_CaptionActive;
		clrBackgroundPaneCaption = minFrameClr.clr_BackgroundMiniFrame_CaptionActive;

		clrTextPaneCaption = minFrameClr.clr_TextCaptionMiniFrame_Active;
	}
	else
	{
		clrButtonPaneCaption = minFrameClr.clr_BackgroundMiniFrame_CaptionDisable;
		clrBackgroundPaneCaption = minFrameClr.clr_BackgroundMiniFrame_CaptionDisable;

		clrTextPaneCaption = minFrameClr.clr_TextCaptionMiniFrame_Disable;
	}

	pDC->FillSolidRect(rectCaption, clrBackgroundPaneCaption);
	pDC->FillSolidRect(rectButtons, clrButtonPaneCaption);

	return clrTextPaneCaption;
}

COLORREF CNPUP_VisualManager::OnFillCaptionBarButton
(
	CDC * pDC,
	CMFCCaptionBar * pBar,
	CRect rect,
	BOOL bIsPressed,
	BOOL bIsHighlighted,
	BOOL bIsDisabled,
	BOOL bHasDropDownArrow,
	BOOL bIsSysButton
)
{
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	ColorTheme::ColorPaneCaptionsAndButton paneCaptionClr = m_currentTheme.m_ColorPaneCaptionsAndButton;

	COLORREF clrTextCaprion = paneCaptionClr.clr_PaneCaption_ButtonsText;
	COLORREF clrCaption = paneCaptionClr.clr_PaneCaption_ButtonsBackground;

	if (bIsDisabled)
	{
		clrTextCaprion = paneCaptionClr.clr_PaneCaption_ButtonsText_Disable;
		clrCaption = paneCaptionClr.clr_PaneCaption_ButtonsBackground_Disable;
	}
	else
	{
		if (bIsPressed)
		{
			clrTextCaprion = paneCaptionClr.clr_PaneCaption_ButtonsText_Pressed;
			clrCaption = paneCaptionClr.clr_PaneCaption_ButtonsBackground_Pressed;
		}
		else
		{
			if (bIsHighlighted)
			{
				clrTextCaprion = paneCaptionClr.clr_PaneCaption_ButtonsText_Highlight;
				clrCaption = paneCaptionClr.clr_PaneCaption_ButtonsBackground_Highlight;
			}
			else
			{
				clrTextCaprion = paneCaptionClr.clr_PaneCaption_ButtonsText;
				clrCaption = paneCaptionClr.clr_PaneCaption_ButtonsBackground;
			}
		}
	}

	pDC->FillSolidRect(rect, clrCaption);

	return clrTextCaprion;
}

void CNPUP_VisualManager::OnDrawCaptionBarInfoArea
(
	CDC * pDC,
	CMFCCaptionBar * pBar,
	CRect rect
)
{
	ColorTheme::ColorPaneCaptionsAndButton paneCaptionClr = m_currentTheme.m_ColorPaneCaptionsAndButton;
	ColorTheme::ParametrsDrawStandard prmDraw = m_currentTheme.m_sParamStandard;

	COLORREF clrCaption = paneCaptionClr.clr_PaneCaption_BackgroundCaption;

	pDC->FillSolidRect(rect, clrCaption);

	if (prmDraw.m_iSizeBorder_CaptionBar_Top > 0
		|| prmDraw.m_iSizeBorder_CaptionBar_Bottom > 0
		|| prmDraw.m_iSizeBorder_CaptionBar_Right > 0
		|| prmDraw.m_iSizeBorder_CaptionBar_Left > 0)
	{
		NVS_DrawRectBorder
		(pDC, rect, RGB(255, 255, 255),
			prmDraw.m_iSizeBorder_CaptionBar_Top,
			prmDraw.m_iSizeBorder_CaptionBar_Bottom,
			prmDraw.m_iSizeBorder_CaptionBar_Left,
			prmDraw.m_iSizeBorder_CaptionBar_Right
		);
	}
}

COLORREF CNPUP_VisualManager::GetCaptionBarTextColor
(
	CMFCCaptionBar * pBar
)
{
	ColorTheme::ColorPaneCaptionsAndButton paneCaptionClr = m_currentTheme.m_ColorPaneCaptionsAndButton;

	COLORREF colorTextBar = paneCaptionClr.clr_PaneCaption_Text;

	return colorTextBar;
}

void CNPUP_VisualManager::OnDrawCaptionBarButtonBorder
(
	CDC * pDC,
	CMFCCaptionBar * pBar,
	CRect rect,
	BOOL bIsPressed,
	BOOL bIsHighlighted,
	BOOL bIsDisabled,
	BOOL bHasDropDownArrow,
	BOOL bIsSysButton
)
{
	ColorTheme::ColorPaneCaptionsAndButton paneCaptionClr = m_currentTheme.m_ColorPaneCaptionsAndButton;

	CRgn rgnMain;
	CRgn rgnExclude;
	COLORREF colorBorder = paneCaptionClr.clr_PaneCaption_ButtonsBorder;

	rgnMain.CreateRectRgn(rect.left, rect.top, rect.right, rect.bottom);
	rgnExclude.CreateRectRgn(rect.left + 1, rect.top + 1, rect.right - 1, rect.bottom - 1);
	rgnMain.CombineRgn(&rgnMain, &rgnExclude, RGN_DIFF);

	if (bIsDisabled)
	{
		colorBorder = paneCaptionClr.clr_PaneCaption_ButtonsBorder_Disable;
	}
	else
	{
		if (bIsPressed)
		{
			colorBorder = paneCaptionClr.clr_PaneCaption_ButtonsBorder_Pressed;
		}
		else
		{
			if (bIsHighlighted)
			{
				colorBorder = paneCaptionClr.clr_PaneCaption_ButtonsBorder_Highlight;
			}
			else
			{
				colorBorder = paneCaptionClr.clr_PaneCaption_ButtonsBorder;
			}
		}
	}

	pDC->SelectClipRgn(&rgnMain);
	pDC->FillSolidRect(rect, colorBorder);
	pDC->SelectClipRgn(NULL);

	rgnMain.DeleteObject();
	rgnExclude.DeleteObject();
}

void CNPUP_VisualManager::OnDrawShowAllMenuItems
(
	CDC * pDC,
	CRect rect,
	CMFCVisualManager::AFX_BUTTON_STATE state
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	COLORREF colorButton = clrGlblElem.clr_ShowAllMenuBackground;

	if (state == ButtonsIsRegular)
	{
		colorButton = clrGlblElem.clr_ShowAllMenuBackground;
	}
	else if (state == ButtonsIsPressed)
	{
		colorButton = clrGlblElem.clr_ShowAllMenuBackground_Pressed;
	}
	else if (state == ButtonsIsHighlighted)
	{
		colorButton = clrGlblElem.clr_ShowAllMenuBackground_Highlight;
	}

	pDC->FillSolidRect(rect, colorButton);
}

void CNPUP_VisualManager::OnDrawMenuResizeBar
(
	CDC * pDC,
	CRect rect,
	int nResizeFlags
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	COLORREF clrBar = clrGlblElem.clr_ResizeBar;

	pDC->FillSolidRect(rect, clrBar);
}

void CNPUP_VisualManager::OnDrawStatusBarPaneBorder
(
	CDC * pDC,
	CMFCStatusBar * pBar,
	CRect rectPane,
	UINT uiID,
	UINT nStyle
)
{
	ColorTheme::ColorStatusBarPane clrStatusBarPane = m_currentTheme.m_ColorStatusBarPane;

	COLORREF clrPane = clrStatusBarPane.clr_StatusBarPane_Background;

	pDC->FillSolidRect(rectPane, clrPane);
}

void CNPUP_VisualManager::OnDrawStatusBarSizeBox
(
	CDC * pDC,
	CMFCStatusBar * pStatBar,
	CRect rectSizeBox
)
{
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	ColorTheme::ColorStatusBarPane &clrStatusBarPane = m_currentTheme.m_ColorStatusBarPane;

	COLORREF colorBox = clrStatusBarPane.clr_StatusBarPane_Background;

	HFONT hFont = CreateFont(
		15,
		0, 0, 0,
		FW_NORMAL,
		FALSE, FALSE, FALSE,
		SYMBOL_CHARSET,
		OUT_DEFAULT_PRECIS,
		CLIP_DEFAULT_PRECIS,
		DEFAULT_QUALITY,
		DEFAULT_PITCH | FF_DONTCARE,
		_T("Marlett")
	);

	HFONT hOldFont = (HFONT)pDC->SelectObject(hFont);

	pDC->FillSolidRect(rectSizeBox, colorBox);
	pDC->SetTextColor(m_currentTheme.m_sParamCMenu.clr_UserColorImageIcon);
	pDC->SetBkMode(TRANSPARENT);
	pDC->TextOut(rectSizeBox.left, rectSizeBox.top + 5, _T("p"), 1);

	pDC->SelectObject(hOldFont);
}

COLORREF CNPUP_VisualManager::GetStatusBarPaneTextColor
(
	CMFCStatusBar * pStatusBar,
	CMFCStatusBarPaneInfo * pPane
)
{
	ColorTheme::ColorStatusBarPane clrStatusBarPane = m_currentTheme.m_ColorStatusBarPane;

	COLORREF clrText = clrStatusBarPane.clr_StatusBarPane_Text;

	return clrText;
}

void CNPUP_VisualManager::OnDrawMenuBorder
(
	CDC * pDC,
	CMFCPopupMenu * pMenu,
	CRect rect
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;

	COLORREF clrBorder = clrGlblElem.clr_BorderMenu;

	if (pMenu->GetParentRibbonElement() != NULL)
	{
		return;
	}
	else
	{
		pDC->FillSolidRect(rect, clrBorder);
	}
}

void CNPUP_VisualManager::OnDrawMenuShadow
(
	CDC * pDC,
	const CRect & rectClient,
	const CRect & rectExclude,
	int nDepth,
	int iMinBrightness,
	int iMaxBrightness,
	CBitmap * pBmpSaveBottom,
	CBitmap * pBmpSaveRight,
	BOOL bRTL
)
{
	ColorTheme::ColorControlElements clrProgressBar = m_currentTheme.m_ColorContrlElem;

	COLORREF clrBar = RGB(255, 0, 0);
	COLORREF clrChunk = RGB(255, 0, 0);

	clrBar = clrProgressBar.clr_ProgressBar_Background;
	clrChunk = clrProgressBar.clr_ProgressChunk_Background;

	pDC->FillSolidRect(rectClient, clrBar); //фон
	pDC->FillSolidRect(rectExclude, clrChunk); //сам бар

}

void CNPUP_VisualManager::OnDrawBarGripper
(
	CDC * pDC,
	CRect rectGripper,
	BOOL bHorz,
	CBasePane * pBar
)
{
	ColorTheme::ColorToolbars toolBarClr = m_currentTheme.m_ColorToolbars;

	COLORREF clrGripper = toolBarClr.clr_ToolbarGripper;

	pDC->FillSolidRect(rectGripper, clrGripper);
}

void CNPUP_VisualManager::OnDrawSeparator
(
	CDC * pDC,
	CBasePane * pBar,
	CRect rect,
	BOOL bIsHoriz
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;

	COLORREF clrSeparator = clrGlblElem.clr_Separator;
	pDC->FillSolidRect(rect, clrSeparator);
}

COLORREF CNPUP_VisualManager::OnDrawMenuLabel
(
	CDC * pDC,
	CRect rect
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;

	COLORREF clrBackground = clrGlblElem.clr_BackgroundLabelMenu;
	COLORREF clrText = clrGlblElem.clr_TextLabelMenu;

	pDC->FillSolidRect(rect, clrBackground);

	return clrText;
}

void CNPUP_VisualManager::OnDrawComboDropButton
(
	CDC * pDC,
	CRect rect,
	BOOL bDisabled,
	BOOL bIsDropped,
	BOOL bIsHighlighted,
	CMFCToolBarComboBoxButton * pButton
)
{
	ColorTheme::ColorMiniFrame minFrameClr = m_currentTheme.m_ColorMiniFrame;
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	CMenuImages currentMenuImages = m_CMIUserImage;

	COLORREF clrButton = minFrameClr.clr_ButtonsElement_Active;

	if (bDisabled)
	{
		clrButton = minFrameClr.clr_ButtonsElement_Disable;
	}
	else
	{
		if (bIsHighlighted)
		{
			clrButton = minFrameClr.clr_ButtonsElement_Highlight;
		}
		else
		{
			if (bIsDropped)
			{
				clrButton = minFrameClr.clr_ButtonsElement_Dropped;
			}
			else
			{
				clrButton = minFrameClr.clr_ButtonsElement_Active;
			}
		}
	}


	pDC->FillSolidRect(rect, clrButton);

	currentMenuImages.Draw(pDC, CMenuImages::IdArrowDown, rect, clrImage);

}

BOOL CNPUP_VisualManager::OnEraseMDIClientArea
(
	CDC * pDC,
	CRect rectClient
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;

	COLORREF clrArea = clrGlblElem.clr_BackgroundClientFrame;

	pDC->FillSolidRect(rectClient, clrArea);

	return TRUE;
}
#pragma endregion

#pragma region AutoHide
void CNPUP_VisualManager::OnFillAutoHideButtonBackground
(
	CDC * pDC,
	CRect rect,
	CMFCAutoHideButton * pButton
)
{
	ColorTheme::ColorTabs tabClr = m_currentTheme.m_ColorTabs;

	COLORREF clrBackground = tabClr.clr_BackgroundTabsHide;

	BOOL isHighlighted = pButton->IsHighlighted();
	BOOL isActive = pButton->IsActive();

	pButton->m_bOverlappingTabs = FALSE;


	DWORD dwAlign = pButton->GetAlignment();

	if (dwAlign == CBRS_ALIGN_LEFT || dwAlign == CBRS_ALIGN_RIGHT)
	{
		if (isHighlighted)
		{
			if (isActive)
			{
				clrBackground = tabClr.clr_BackgroundTabsHide_Active;
			}
			else
			{
				clrBackground = tabClr.clr_BackgroundTabsHide_Highlight;
			}
		}
		else
		{
			if (isActive)
			{
				clrBackground = tabClr.clr_BackgroundTabsHide_Active;
			}
			else
			{
				clrBackground = tabClr.clr_BackgroundTabsHide_Disable;
			}
		}
	}
	else if (dwAlign == CBRS_ALIGN_TOP || dwAlign == CBRS_ALIGN_BOTTOM)
	{
		if (isHighlighted)
		{
			clrBackground = tabClr.clr_BackgroundTabsHide_Highlight;
		}
		else
		{
			clrBackground = tabClr.clr_BackgroundTabsHide_Disable;
		}

	}

	pDC->FillSolidRect(rect, clrBackground);

}

void CNPUP_VisualManager::OnDrawAutoHideButtonBorder
(
	CDC * pDC,
	CRect rectBounds,
	CRect rectBorderSize,
	CMFCAutoHideButton * pButton
)
{
	ColorTheme::ColorTabs tabClr = m_currentTheme.m_ColorTabs;
	ColorTheme::ParametrsDrawStandard prmStndr = m_currentTheme.m_sParamStandard;

	COLORREF colorBorder = tabClr.clr_BorderTabsHide;
	CDC memDC;

	BOOL bIsActive = pButton->IsActive();
	BOOL bIsHighlighted = pButton->IsHighlighted();

	DWORD dwAlign = pButton->GetAlignment();

	if (dwAlign == CBRS_ALIGN_LEFT || dwAlign == CBRS_ALIGN_RIGHT)
	{
		if (bIsHighlighted)
		{
			if (bIsActive)
			{
				colorBorder = tabClr.clr_BorderTabsHide_Active;
			}
			else
			{
				colorBorder = tabClr.clr_BorderTabsHide_Highlight;
			}
		}
		else
		{
			if (bIsActive)
			{
				colorBorder = tabClr.clr_BorderTabsHide_Active;
			}
			else
			{
				colorBorder = tabClr.clr_BorderTabsHide_Disable;
			}
		}
	}
	else if (dwAlign == CBRS_ALIGN_TOP || dwAlign == CBRS_ALIGN_BOTTOM)
	{
		if (bIsHighlighted)
		{
			colorBorder = tabClr.clr_BorderTabsHide_Highlight;
		}
		else
		{
			colorBorder = tabClr.clr_BorderTabsHide_Disable;
		}
	}

	if (prmStndr.m_bUnderlineTabs_AutoHideTab)
	{
		if (prmStndr.m_iSizeBorder_AutoHideTab_Top > 0
			|| prmStndr.m_iSizeBorder_AutoHideTab_Bottom > 0
			|| prmStndr.m_iSizeBorder_AutoHideTab_Right > 0
			|| prmStndr.m_iSizeBorder_AutoHideTab_Left > 0)
		{

			if (dwAlign == CBRS_ALIGN_LEFT || dwAlign == CBRS_ALIGN_RIGHT)
			{
				if (prmStndr.m_bBorderUnderLine_AutoHideTab)
				{
					NVS_DrawRectBorder
					(
						pDC, rectBounds, colorBorder,
						1,
						1,
						1,
						1
					);
				}

				rectBounds.bottom -= prmStndr.m_iSizeUnderlineTabs_AutoHideTab;
				rectBounds.top += prmStndr.m_iSizeUnderlineTabs_AutoHideTab;

				NVS_DrawRectBorder
				(
					pDC, rectBounds, colorBorder,
					0,
					0,
					prmStndr.m_iSizeBorder_AutoHideTab_Bottom,
					prmStndr.m_iSizeBorder_AutoHideTab_Top
				);
			}
			else if (dwAlign == CBRS_ALIGN_TOP || dwAlign == CBRS_ALIGN_BOTTOM)
			{
				if (prmStndr.m_bBorderUnderLine_AutoHideTab)
				{
					NVS_DrawRectBorder
					(
						pDC, rectBounds, colorBorder,
						1,
						1,
						1,
						1
					);
				}

				rectBounds.left += prmStndr.m_iSizeUnderlineTabs_AutoHideTab;
				rectBounds.right -= prmStndr.m_iSizeUnderlineTabs_AutoHideTab;

				NVS_DrawRectBorder
				(
					pDC, rectBounds, colorBorder,
					prmStndr.m_iSizeBorder_AutoHideTab_Top,
					prmStndr.m_iSizeBorder_AutoHideTab_Bottom,
					0,
					0
				);
			}
		}
	}
	else
	{
		if (prmStndr.m_iSizeBorder_AutoHideTab_Top > 0
			|| prmStndr.m_iSizeBorder_AutoHideTab_Bottom > 0
			|| prmStndr.m_iSizeBorder_AutoHideTab_Right > 0
			|| prmStndr.m_iSizeBorder_AutoHideTab_Left > 0)
		{

			if (dwAlign == CBRS_ALIGN_LEFT || dwAlign == CBRS_ALIGN_RIGHT)
			{
				NVS_DrawRectBorder
				(
					pDC, rectBounds, colorBorder,
					prmStndr.m_iSizeBorder_AutoHideTab_Left,
					prmStndr.m_iSizeBorder_AutoHideTab_Right,
					prmStndr.m_iSizeBorder_AutoHideTab_Bottom,
					prmStndr.m_iSizeBorder_AutoHideTab_Top
				);
			}
			else if (dwAlign == CBRS_ALIGN_TOP || dwAlign == CBRS_ALIGN_BOTTOM)
			{
				NVS_DrawRectBorder
				(
					pDC, rectBounds, colorBorder,
					prmStndr.m_iSizeBorder_AutoHideTab_Top,
					prmStndr.m_iSizeBorder_AutoHideTab_Bottom,
					prmStndr.m_iSizeBorder_AutoHideTab_Left,
					prmStndr.m_iSizeBorder_AutoHideTab_Right
				);
			}
		}
	}

}

COLORREF CNPUP_VisualManager::GetAutoHideButtonTextColor
(
	CMFCAutoHideButton * pButton
)
{
	ColorTheme::ColorTabs tabClr = m_currentTheme.m_ColorTabs;

	COLORREF clrText = tabClr.clr_TextTabsHide;

	BOOL bIsActive = pButton->IsActive();
	BOOL bIsHighlighted = pButton->IsHighlighted();

	DWORD dwAlign = pButton->GetAlignment();

	if (dwAlign == CBRS_ALIGN_LEFT || dwAlign == CBRS_ALIGN_RIGHT)
	{
		if (bIsHighlighted)
		{
			if (bIsActive)
			{
				clrText = tabClr.clr_TextTabsHide_Active;
			}
			else
			{
				clrText = tabClr.clr_TextTabsHide_Highlight;
			}
		}
		else
		{
			if (bIsActive)
			{
				clrText = tabClr.clr_TextTabsHide_Active;
			}
			else
			{
				clrText = tabClr.clr_TextTabsHide_Disable;
			}
		}
	}
	else if (dwAlign == CBRS_ALIGN_TOP || dwAlign == CBRS_ALIGN_BOTTOM)
	{
		if (bIsHighlighted)
		{
			clrText = tabClr.clr_TextTabsHide_Highlight;
		}
		else
		{
			clrText = tabClr.clr_TextTabsHide_Disable;
		}
	}

	return clrText;
}

#pragma endregion

#pragma region MiniFrame
COLORREF CNPUP_VisualManager::OnFillMiniFrameCaption
(
	CDC * pDC,
	CRect rectCaption,
	CPaneFrameWnd * pFrameWnd,
	BOOL bActive
)
{
	ColorTheme::ColorMiniFrame minFrameClr = m_currentTheme.m_ColorMiniFrame;

	COLORREF clrCaption = minFrameClr.clr_BackgroundMiniFrame;
	COLORREF clrText = minFrameClr.clr_TextCaptionMiniFrame_Active;

	if (bActive)
	{
		clrCaption = minFrameClr.clr_BackgroundMiniFrame_CaptionActive;
		clrText = minFrameClr.clr_TextCaptionMiniFrame_Active;
	}
	else
	{
		clrCaption = minFrameClr.clr_BackgroundMiniFrame_CaptionDisable;
		clrText = minFrameClr.clr_TextCaptionMiniFrame_Disable;
	}

	pDC->FillSolidRect(rectCaption, clrCaption);
	return clrText;
}

void CNPUP_VisualManager::OnDrawMiniFrameBorder
(
	CDC * pDC,
	CPaneFrameWnd * pFrameWnd,
	CRect rectBorder,
	CRect rectBorderSize
)
{
	ColorTheme::ColorMiniFrame minFrameClr = m_currentTheme.m_ColorMiniFrame;
	COLORREF colorBorder = minFrameClr.clr_BorderMiniFrame_Client;

	pDC->FillSolidRect(rectBorder, colorBorder);
}

void CNPUP_VisualManager::OnDrawCaptionButton
(
	CDC * pDC,
	CMFCCaptionButton * pButton,
	BOOL bActive,
	BOOL bHorz,
	BOOL bMaximized,
	BOOL bDisabled,
	int nImageID
)
{
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	CMenuImages currentMenuImages = m_CMIUserImage;
	ColorTheme::ColorMiniFrame minFrameClr = m_currentTheme.m_ColorMiniFrame;

	COLORREF clrButton = minFrameClr.clr_ButtonMiniFrame;

	CRect rectButton = pButton->GetRect();

	BOOL bFocused = pButton->m_bFocused;

	if (bActive)
	{
		if (bFocused)
		{
			clrButton = minFrameClr.clr_ButtonMiniFrame_Active;
		}
		else
		{
			clrButton = minFrameClr.clr_ButtonMiniFrame_Highlight;
		}
	}
	else
	{
		if (bFocused)
		{
			clrButton = minFrameClr.clr_ButtonMiniFrame_Active;
		}
		else
		{
			clrButton = minFrameClr.clr_ButtonMiniFrame_Disable;
		}
	}

	pDC->FillSolidRect(rectButton, clrButton);
	currentMenuImages.Draw(pDC, pButton->GetIconID(bHorz, bMaximized), rectButton, clrImage);
}

void CNPUP_VisualManager::OnDrawPaneDivider
(
	CDC * pDC,
	CPaneDivider * pSlider,
	CRect rect,
	BOOL bAutoHideMode
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	COLORREF clrDivider = clrGlblElem.clr_DividerFrame;

	pDC->FillSolidRect(rect, clrDivider);
}

#pragma endregion

#pragma region PropertyGrig
COLORREF CNPUP_VisualManager::GetPropertyGridGroupColor
(
	CMFCPropertyGridCtrl * pPropList
)
{
	ColorTheme::ColorPropertyGrid propGridColor = m_currentTheme.m_ColorPropertyGrid;

	COLORREF returnColor = propGridColor.clr_PropGrid_Line;


	return returnColor;
}

COLORREF CNPUP_VisualManager::GetPropertyGridGroupTextColor
(
	CMFCPropertyGridCtrl * pPropList
)
{
	ColorTheme::ColorPropertyGrid propGridColor = m_currentTheme.m_ColorPropertyGrid;

	COLORREF returnColor = propGridColor.clr_PropGrid_Line;

	return returnColor;
}

void CNPUP_VisualManager::OnDrawExpandingBox
(
	CDC * pDC,
	CRect rect,
	BOOL bIsOpened,
	COLORREF colorBox
)
{
	ColorTheme::ColorMiniFrame minFrameClr = m_currentTheme.m_ColorMiniFrame;
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	CMenuImages currentMenuImages = m_CMIUserImage;

	CMenuImages::IMAGES_IDS cImageBox = CMenuImages::IdCloseBold;
	COLORREF clrBackground = minFrameClr.clr_ButtonsElement;

	if (bIsOpened)
	{
		clrBackground = minFrameClr.clr_ButtonsElement_Dropped;

		cImageBox = CMenuImages::IdArrowDown;
	}
	else
	{
		clrBackground = minFrameClr.clr_ButtonsElement_NoDropped;

		cImageBox = CMenuImages::IdArrowRight;
	}


	pDC->FillSolidRect(rect, clrBackground);

	currentMenuImages.Draw(pDC, cImageBox, rect, clrImage);

}

#pragma endregion

#pragma region Tabs
void CNPUP_VisualManager::GetTabFrameColors
(
	const CMFCBaseTabCtrl * pTabWnd,
	COLORREF & clrDark,
	COLORREF & clrBlack,
	COLORREF & clrHighlight,
	COLORREF & clrFace,
	COLORREF & clrDarkShadow,
	COLORREF & clrLight,
	CBrush *& pbrFace,
	CBrush *& pbrBlack
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;

	clrDark = clrGlblElem.clr_BorderFrame;
	clrBlack = clrGlblElem.clr_BorderFrame;
	clrHighlight = clrGlblElem.clr_BorderFrame;
	clrFace = clrGlblElem.clr_BorderFrame;
	clrDarkShadow = clrGlblElem.clr_BorderFrame;
	clrLight = clrGlblElem.clr_BorderFrame;

	pbrFace = new CBrush();
	pbrBlack = new CBrush();

	pbrFace->CreateSolidBrush(clrGlblElem.clr_BorderFrame);
	pbrBlack->CreateSolidBrush(clrGlblElem.clr_BorderFrame);
}

void CNPUP_VisualManager::OnEraseTabsArea
(
	CDC * pDC,
	CRect rect,
	const CMFCBaseTabCtrl * pTabWnd
)
{
	ColorTheme::ColorTabs tabClr = m_currentTheme.m_ColorTabs;

	COLORREF clrTabsBackground = tabClr.clr_BackgroundFrameDockMDI;

	pDC->FillSolidRect(rect, clrTabsBackground);
}

BOOL CNPUP_VisualManager::OnEraseTabsFrame
(
	CDC * pDC,
	CRect rect,
	const CMFCBaseTabCtrl * pTabWnd
)
{
	ColorTheme::ColorMiniFrame minFrameClr = m_currentTheme.m_ColorMiniFrame;

	COLORREF clrFrame = minFrameClr.clr_BorderMiniFrame_Client;

	pDC->FillSolidRect(rect, clrFrame);

	return TRUE;
}

void CNPUP_VisualManager::OnEraseTabsButton
(
	CDC * pDC,
	CRect rect,
	CMFCButton * pButton,
	CMFCBaseTabCtrl * pWndTab
)
{
	ColorTheme::ColorMiniFrame minFrameClr = m_currentTheme.m_ColorMiniFrame;

	COLORREF clrButton = minFrameClr.clr_ButtonsElement;

	BOOL bIsHighlighted = pButton->IsHighlighted();
	BOOL bIsPressed = pButton->IsPressed();

	if (bIsHighlighted)
	{
		clrButton = minFrameClr.clr_ButtonMiniFrame_Highlight;
	}
	else
	{
		if (bIsPressed)
		{
			clrButton = minFrameClr.clr_TextCaptionMiniFrame_Active;
		}
		else
		{
			clrButton = minFrameClr.clr_TextCaptionMiniFrame_Disable;
		}
	}

	pDC->FillSolidRect(rect, clrButton);
}

void CNPUP_VisualManager::OnDrawTabsButtonBorder
(
	CDC * pDC,
	CRect & rect,
	CMFCButton * pButton,
	UINT uiState,
	CMFCBaseTabCtrl * pWndTab
)
{
	ColorTheme::ColorMiniFrame minFrameClr = m_currentTheme.m_ColorMiniFrame;

	COLORREF clrButton = minFrameClr.clr_ButtonsBorderElement;

	BOOL bIsHighlighted = pButton->IsHighlighted();
	BOOL bIsPressed = pButton->IsPressed();

	if (bIsHighlighted)
	{
		clrButton = minFrameClr.clr_ButtonsBorderElement_Highlight;
	}
	else
	{
		if (bIsPressed)
		{
			clrButton = minFrameClr.clr_ButtonsBorderElement_Active;
		}
		else
		{
			clrButton = minFrameClr.clr_ButtonsBorderElement_Disable;
		}
	}

	NVS_DrawRectBorder(pDC, rect, clrButton, 1, 1, 1, 1);
}

void CNPUP_VisualManager::OnDrawTab
(
	CDC * pDC,
	CRect rectTab,
	int iTab,
	BOOL bIsActive,
	const CMFCBaseTabCtrl * pTabWnd
)
{
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	CMenuImages currentMenuImages = m_CMIUserImage;

	ColorTheme::ColorTabs tabClr = m_currentTheme.m_ColorTabs;
	ColorTheme::ParametrsDrawStandard prmStdr = m_currentTheme.m_sParamStandard;

	COLORREF clrTab = tabClr.clr_BackgroundTabsMDI;
	COLORREF clrBorderTab = tabClr.clr_BorderTabsMDI;
	COLORREF clrText = tabClr.clr_TextTabsMDI;
	CString strLabel;

	pTabWnd->GetTabLabel(iTab, strLabel);

	CPoint ptCursor;
	::GetCursorPos(&ptCursor);
	::ScreenToClient(pTabWnd->GetSafeHwnd(), &ptCursor);
	BOOL bIsHover = rectTab.PtInRect(ptCursor);

	if (bIsHover)
	{
		if (bIsActive)
		{
			clrTab = tabClr.clr_BackgroundTabsMDI_Active;
			clrBorderTab = tabClr.clr_BorderTabsMDI_Active;
			clrText = tabClr.clr_TextTabsMDI_Active;

		}
		else
		{
			clrTab = tabClr.clr_BackgroundTabsMDI_Highlight;
			clrBorderTab = tabClr.clr_BorderTabsMDI_Highlight;
			clrText = tabClr.clr_TextTabsMDI_Highlight;
		}
	}
	else
	{
		if (bIsActive)
		{
			clrTab = tabClr.clr_BackgroundTabsMDI_Active;
			clrBorderTab = tabClr.clr_BorderTabsMDI_Active;
			clrText = tabClr.clr_TextTabsMDI_Active;
		}
		else
		{
			clrTab = tabClr.clr_BackgroundTabsMDI_Disable;
			clrBorderTab = tabClr.clr_BorderTabsMDI_Disable;
			clrText = tabClr.clr_TextTabsMDI_Disable;
		}
	}

	pDC->FillSolidRect(rectTab, clrTab);

	int nOldBkMode = pDC->SetBkMode(TRANSPARENT);
	COLORREF clrOldText = pDC->SetTextColor(clrText);
	CFont* pOldFont = pDC->SelectObject(&afxGlobalData.fontRegular);

	CRect rectText = rectTab;
	rectText.DeflateRect(2, 0);

	pDC->DrawText(strLabel, rectText, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

	if (pTabWnd->IsActiveTabCloseButton() == TRUE
		&& bIsActive == TRUE)
	{
		COLORREF clrBox = tabClr.clr_ButtonTabsMDI;
		COLORREF clrBorder = tabClr.clr_BorderButtonTabsMDI;
		CRect rectBox = rectTab;

		int nBoxSize = rectTab.Height();

		rectBox.left = rectBox.right - nBoxSize + 2;
		rectBox.right -= 2;
		rectBox.top += 3;
		rectBox.bottom -= 1;

		if (pTabWnd->IsTabCloseButtonHighlighted())
		{
			clrBox = tabClr.clr_ButtonTabsMDI_Highlight;
			clrBorder = tabClr.clr_BorderButtonTabsMDI_Highlight;
		}
		else
		{
			if (bIsActive)
			{
				clrBox = tabClr.clr_ButtonTabsMDI_Active;
				clrBorder = tabClr.clr_BorderButtonTabsMDI_Active;
			}
			else
			{
				clrBox = tabClr.clr_ButtonTabsMDI_Disable;
				clrBorder = tabClr.clr_BorderButtonTabsMDI_Disable;
			}
		}

		pDC->FillSolidRect(rectBox, clrBox);

		if (prmStdr.m_iSizeBorder_MDIButtonTab_Top > 0
			|| prmStdr.m_iSizeBorder_MDIButtonTab_Bottom > 0
			|| prmStdr.m_iSizeBorder_MDIButtonTab_Right > 0
			|| prmStdr.m_iSizeBorder_MDIButtonTab_Left > 0)
		{
			NVS_DrawRectBorder
			(
				pDC, rectBox, clrBox,
				prmStdr.m_iSizeBorder_MDIButtonTab_Top,
				prmStdr.m_iSizeBorder_MDIButtonTab_Bottom,
				prmStdr.m_iSizeBorder_MDIButtonTab_Left,
				prmStdr.m_iSizeBorder_MDIButtonTab_Right
			);
		}


		rectBox.bottom -= 2;

		currentMenuImages.Draw(pDC, CMenuImages::IdClose, rectBox, clrImage);

	}

	if (prmStdr.m_bUnderlineTabs_MDITab)
	{
		if (prmStdr.m_iSizeBorder_MDITab_Top > 0
			|| prmStdr.m_iSizeBorder_MDITab_Bottom > 0
			|| prmStdr.m_iSizeBorder_MDITab_Right > 0
			|| prmStdr.m_iSizeBorder_MDITab_Left > 0)
		{
			if (prmStdr.m_bBorderUnderLine_MDITab)
			{
				NVS_DrawRectBorder
				(
					pDC, rectTab, clrBorderTab,
					1,
					1,
					1,
					1
				);
			}

			rectTab.left += prmStdr.m_iSizeUnderlineTabs_MDITab;
			rectTab.right -= prmStdr.m_iSizeUnderlineTabs_MDITab;

			NVS_DrawRectBorder
			(
				pDC, rectTab, clrBorderTab,
				prmStdr.m_iSizeBorder_MDITab_Top,
				prmStdr.m_iSizeBorder_MDITab_Bottom,
				0,
				0
			);
		}
	}
	else
	{
		if (prmStdr.m_iSizeBorder_MDITab_Top > 0
			|| prmStdr.m_iSizeBorder_MDITab_Bottom > 0
			|| prmStdr.m_iSizeBorder_MDITab_Right > 0
			|| prmStdr.m_iSizeBorder_MDITab_Left > 0)
		{
			NVS_DrawRectBorder
			(
				pDC, rectTab, clrBorderTab,
				prmStdr.m_iSizeBorder_MDITab_Top,
				prmStdr.m_iSizeBorder_MDITab_Bottom,
				prmStdr.m_iSizeBorder_MDITab_Left,
				prmStdr.m_iSizeBorder_MDITab_Right
			);
		}
	}


	pDC->SelectObject(pOldFont);
	pDC->SetTextColor(clrOldText);
	pDC->SetBkMode(nOldBkMode);


}

void CNPUP_VisualManager::OnDrawTabCloseButton
(
	CDC * pDC,
	CRect rect,
	const CMFCBaseTabCtrl * pTabWnd,
	BOOL bIsHighlighted,
	BOOL bIsPressed,
	BOOL bIsDisabled
)
{
	// метод не используется

	ColorTheme::ColorTabs tabClr = m_currentTheme.m_ColorTabs;
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	CMenuImages currentMenuImages = m_CMIUserImage;

	COLORREF clrBox = RGB(0, 32, 32);

	if (bIsHighlighted)
	{
		clrBox = tabClr.clr_ButtonTabsMDI_Highlight;
	}
	else
	{
		clrBox = tabClr.clr_ButtonTabsMDI_Active;
	}

	pDC->FillSolidRect(rect, clrBox);

	currentMenuImages.Draw(pDC, CMenuImages::IdClose, rect, clrImage);

}

void CNPUP_VisualManager::OnDrawTabContent
(
	CDC * pDC,
	CRect rectTab,
	int iTab,
	BOOL bIsActive,
	const CMFCBaseTabCtrl * pTabWnd,
	COLORREF clrText
)
{
	// метод не используется

	COLORREF colorTab = RGB(32, 32, 255);;
	COLORREF colorText = RGB(255, 255, 255);

	clrText = colorText;
	pDC->FillSolidRect(rectTab, colorTab);
}

void CNPUP_VisualManager::OnDrawTabResizeBar
(
	CDC * pDC,
	CMFCBaseTabCtrl * pWndTab,
	BOOL bIsVert,
	CRect rect,
	CBrush * pbrFace,
	CPen * pPen
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	COLORREF clrBar = clrGlblElem.clr_DividerFrame;

	if (bIsVert)
	{
		clrBar = clrGlblElem.clr_DividerFrame;
	}
	else
	{
		clrBar = clrGlblElem.clr_DividerFrame;
	}

	pDC->FillSolidRect(rect, clrBar);
}

void CNPUP_VisualManager::OnFillTab
(
	CDC * pDC,
	CRect rectFill,
	CBrush * pbrFill,
	int iTab,
	BOOL bIsActive,
	const CMFCBaseTabCtrl * pTabWnd
)
{
	ColorTheme::ColorTabs tabClr = m_currentTheme.m_ColorTabs;
	ColorTheme::ParametrsDrawStandard prmDrawStandard = m_currentTheme.m_sParamStandard;

	COLORREF clrBackgroundTab = tabClr.clr_BackgroundTabsMDI;

	if (!prmDrawStandard.m_bStandartDrawingTab)
	{
		return;
	}

	if (bIsActive)
	{
		clrBackgroundTab = tabClr.clr_BackgroundTabsMDI_Active;
	}
	else
	{
		clrBackgroundTab = tabClr.clr_BackgroundTabsMDI_Disable;
	}

	pDC->FillSolidRect(rectFill, clrBackgroundTab);

}

#pragma endregion

#pragma region OutlookBar
void CNPUP_VisualManager::OnFillOutlookPageButton
(
	CDC * pDC,
	const CRect & rect,
	BOOL bIsHighlighted,
	BOOL bIsPressed,
	COLORREF & clrText
)
{
	ColorTheme::ColorOutlookBar outlookBarClr = m_currentTheme.m_ColorOutlookBar;

	COLORREF clrTextButton = outlookBarClr.clr_OutlookBarButtonText;
	COLORREF clrBackground = outlookBarClr.clr_OutlookBarButtonBackground;

	if (bIsHighlighted)
	{
		clrTextButton = outlookBarClr.clr_OutlookBarButtonText_Highlight;
		clrBackground = outlookBarClr.clr_OutlookBarButtonBackground_Highlight;
	}
	else
	{
		if (bIsPressed)
		{
			clrTextButton = outlookBarClr.clr_OutlookBarButtonText_Pressed;
			clrBackground = outlookBarClr.clr_OutlookBarButtonBackground_Pressed;
		}
		else
		{
			clrTextButton = outlookBarClr.clr_OutlookBarButtonText;
			clrBackground = outlookBarClr.clr_OutlookBarButtonBackground;
		}
	}

	clrText = clrTextButton;

	pDC->FillSolidRect(rect, clrBackground);
}

void CNPUP_VisualManager::OnDrawOutlookPageButtonBorder
(
	CDC * pDC,
	CRect & rectBtn,
	BOOL bIsHighlighted,
	BOOL bIsPressed
)
{
	ColorTheme::ColorOutlookBar outlookBarClr = m_currentTheme.m_ColorOutlookBar;

	COLORREF clrBorder = outlookBarClr.clr_OutlookBarButtonBorder;

	if (bIsHighlighted)
	{
		clrBorder = outlookBarClr.clr_OutlookBarButtonBorder_Highlight;
	}
	else
	{
		if (bIsPressed)
		{
			clrBorder = outlookBarClr.clr_OutlookBarButtonBorder_Pressed;
		}
		else
		{
			clrBorder = outlookBarClr.clr_OutlookBarButtonBorder;
		}
	}

	NVS_DrawRectBorder(pDC, rectBtn, clrBorder, 1, 2, 3, 4);

}

void CNPUP_VisualManager::OnDrawOutlookBarSplitter
(
	CDC * pDC,
	CRect rectSplitter
)
{
	ColorTheme::ColorOutlookBar outlookBarClr = m_currentTheme.m_ColorOutlookBar;
	COLORREF clrBorder = outlookBarClr.clr_OutlookBarSeparator;

	pDC->FillSolidRect(rectSplitter, clrBorder);
}

void CNPUP_VisualManager::OnFillOutlookBarCaption
(
	CDC * pDC,
	CRect rectCaption,
	COLORREF & clrText
)
{
	ColorTheme::ColorOutlookBar outlookBarClr = m_currentTheme.m_ColorOutlookBar;
	COLORREF clrCaption = outlookBarClr.clr_OutlookBarCaptionBackground;

	clrText = outlookBarClr.clr_OutlookBarCaptionText;
	pDC->FillSolidRect(rectCaption, clrCaption);
}
#pragma endregion

#pragma region ToolbarAndMenu
void CNPUP_VisualManager::OnDrawFloatingToolbarBorder
(
	CDC * pDC,
	CMFCBaseToolBar * pToolBar,
	CRect rectBorder,
	CRect rectBorderSize
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	COLORREF colorBorder = clrGlblElem.clr_BorderFrame;

	pDC->FillSolidRect(rectBorder, colorBorder);
}

void CNPUP_VisualManager::OnDrawButtonBorder
(
	CDC * pDC,
	CMFCToolBarButton * pButton,
	CRect rect,
	CMFCVisualManager::AFX_BUTTON_STATE state
)
{
	ColorTheme::ColorToolbars toolBarClr = m_currentTheme.m_ColorToolbars;
	COLORREF colorButton = toolBarClr.clr_BorderIconToolbar;

	if (state == CMFCVisualManager::ButtonsIsRegular)
	{
		colorButton = toolBarClr.clr_BorderIconToolbar;
	}
	else if (state == CMFCVisualManager::ButtonsIsPressed)
	{
		colorButton = toolBarClr.clr_BorderIconToolbar_Pressed;
	}
	else if (state == CMFCVisualManager::ButtonsIsHighlighted)
	{
		colorButton = toolBarClr.clr_BorderIconToolbar_Highlight;
	}

	NVS_DrawRectBorder(pDC, rect, colorButton, 1, 1, 1, 1);

}

void CNPUP_VisualManager::OnHighlightRarelyUsedMenuItems
(
	CDC * pDC,
	CRect rectRarelyUsed
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;

	COLORREF colorMenuItem = clrGlblElem.clr_ImportantMenu;

	pDC->FillSolidRect(rectRarelyUsed, colorMenuItem);
}


void CNPUP_VisualManager::OnHighlightMenuItem
(
	CDC * pDC,
	CMFCToolBarMenuButton * pButton,
	CRect rect,
	COLORREF & clrText
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	COLORREF colorBackground = clrGlblElem.clr_SelectedMenu;
	COLORREF colorBorder = clrGlblElem.clr_BorderMenu;

	clrText = clrGlblElem.clr_TextMenu_Selected;

	pDC->FillSolidRect(rect, colorBackground);
	NVS_DrawRectBorder(pDC, rect, colorBorder, 1, 1, 1, 1);
}


void CNPUP_VisualManager::OnDrawMenuCheck
(
	CDC * pDC,
	CMFCToolBarMenuButton * pButton,
	CRect rect,
	BOOL bHighlight,
	BOOL bIsRadio
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	CMenuImages currentMenuImages = m_CMIUserImage;

	COLORREF colorButton = clrGlblElem.clr_CheckMenu;

	pDC->FillSolidRect(rect, colorButton);

	if (bIsRadio)
	{
		currentMenuImages.Draw(pDC, CMenuImages::IdRadio, rect, clrImage);
	}
	else
	{
		currentMenuImages.Draw(pDC, CMenuImages::IdCheck, rect, clrImage);
	}

}

void CNPUP_VisualManager::OnFillButtonInterior
(
	CDC * pDC,
	CMFCToolBarButton * pButton,
	CRect rect,
	CMFCVisualManager::AFX_BUTTON_STATE state
)
{
	ColorTheme::ColorToolbars toolBarClr = m_currentTheme.m_ColorToolbars;
	COLORREF colorBackGround = toolBarClr.clr_BackgroundIconToolbar;

	if (state == CMFCVisualManager::ButtonsIsRegular)
	{
		colorBackGround = toolBarClr.clr_BackgroundIconToolbar;
	}
	else if (state == CMFCVisualManager::ButtonsIsPressed)
	{
		colorBackGround = toolBarClr.clr_BackgroundIconToolbar_Pressed;
	}
	else if (state == CMFCVisualManager::ButtonsIsHighlighted)
	{
		colorBackGround = toolBarClr.clr_BackgroundIconToolbar_Highlight;
	}

	pDC->FillSolidRect(rect, colorBackGround);
}

BOOL CNPUP_VisualManager::IsToolbarRoundShape(CMFCToolBar * pToolBar)
{
	return FALSE;
}
#pragma endregion

#pragma region TaskPane

#pragma endregion

#pragma region Other
COLORREF CNPUP_VisualManager::OnFillCommandsListBackground
(
	CDC * pDC,
	CRect rect,
	BOOL bIsSelected
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;

	COLORREF clrBackground = clrGlblElem.clr_CommandListBackground;
	COLORREF clrTextCommand = clrGlblElem.clr_TextCommandList;

	if (bIsSelected)
	{
		clrBackground = clrGlblElem.clr_CommandListBackground_Selected;
		clrTextCommand = clrGlblElem.clr_TextCommandList_Selected;
	}
	else
	{
		clrBackground = clrGlblElem.clr_CommandListBackground;
		clrTextCommand = clrGlblElem.clr_TextCommandList;
	}

	pDC->FillSolidRect(rect, clrBackground);

	return clrTextCommand;
}

void CNPUP_VisualManager::OnDrawCheckBoxEx
(
	CDC * pDC,
	CRect rect,
	int nState,
	BOOL bHighlighted,
	BOOL bPressed,
	BOOL bEnabled
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	CMenuImages currentMenuImages = m_CMIUserImage;

	COLORREF colorBackground = clrGlblElem.clr_CheckBoxBackground;
	COLORREF colorFrame = clrGlblElem.clr_CheckBoxBorder;

	if (bPressed)
	{
		colorBackground = clrGlblElem.clr_CheckBoxBackground_Pressed;
		colorFrame = clrGlblElem.clr_CheckBoxBorder_Pressed;
	}
	else
	{
		if (bHighlighted)
		{
			colorBackground = clrGlblElem.clr_CheckBoxBackground_Highlight;
			colorFrame = clrGlblElem.clr_CheckBoxBorder_Highlight;
		}
		else
		{
			if (bEnabled)
			{
				colorBackground = clrGlblElem.clr_CheckBoxBackground_Enable;
				colorFrame = clrGlblElem.clr_CheckBoxBorder_Enable;
			}
			else
			{
				colorBackground = clrGlblElem.clr_CheckBoxBackground_Disable;
				colorFrame = clrGlblElem.clr_CheckBoxBorder_Disable;
			}
		}
	}

	pDC->FillSolidRect(rect, colorBackground);

	NVS_DrawRectBorder(pDC, rect, colorFrame, 1, 1, 1, 1);

	if (nState)
	{
		currentMenuImages.Draw(pDC, CMenuImages::IdCheck, rect, clrImage);
	}

}

void CNPUP_VisualManager::OnDrawSpinButtons
(
	CDC * pDC,
	CRect rectSpin,
	int nState,
	BOOL bOrientation,
	CMFCSpinButtonCtrl * pSpinCtrl
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	CMenuImages currentMenuImages = m_CMIUserImage;

	COLORREF clrBackgroundUp = clrGlblElem.clr_SpinButtonsBackground_Up;
	COLORREF clrBackgroundDown = clrGlblElem.clr_SpinButtonsBackground_Down;
	COLORREF clrFrameUP = clrGlblElem.clr_SpinButtonsBorder_Up;
	COLORREF clrFrameDown = clrGlblElem.clr_SpinButtonsBorder_Down;

	CRect spinRectUP;
	CRect spinRectDown;

	CRgn rgnMainUP;
	CRgn rgnMainDown;
	CRgn rgnExcludeUP;
	CRgn rgnExcludeDown;

	int sizeFrame = 1;

	spinRectUP.CopyRect(rectSpin);
	spinRectDown.CopyRect(rectSpin);

	if (nState == AFX_SPIN_PRESSEDUP)
	{
		clrBackgroundUp = clrGlblElem.clr_SpinButtonsBackgroundPressed_Up;
		clrFrameUP = clrGlblElem.clr_SpinButtonsBorderPressed_Up;
	}
	else if (nState == AFX_SPIN_PRESSEDDOWN)
	{
		clrBackgroundDown = clrGlblElem.clr_SpinButtonsBackgroundPressed_Down;
		clrFrameDown = clrGlblElem.clr_SpinButtonsBorderPressed_Down;
	}
	else if (nState == AFX_SPIN_HIGHLIGHTEDUP)
	{
		clrBackgroundUp = clrGlblElem.clr_SpinButtonsBackgroundHighlight_Up;
		clrFrameUP = clrGlblElem.clr_SpinButtonsBorderHighlight_Up;
	}
	else if (nState == AFX_SPIN_HIGHLIGHTEDDOWN)
	{
		clrBackgroundDown = clrGlblElem.clr_SpinButtonsBackgroundHighlight_Down;
		clrFrameDown = clrGlblElem.clr_SpinButtonsBorderHighlight_Down;
	}
	else if (nState == AFX_SPIN_DISABLED)
	{
		clrBackgroundUp = clrGlblElem.clr_SpinButtonsBackgroundDisable_Up;
		clrBackgroundDown = clrGlblElem.clr_SpinButtonsBackgroundDisable_Down;
		clrFrameUP = clrGlblElem.clr_SpinButtonsBorderDisable_Up;
		clrFrameDown = clrGlblElem.clr_SpinButtonsBorderDisable_Down;
	}

	if (bOrientation)
	{
		pDC->FillSolidRect(spinRectUP, clrBackgroundUp);

		currentMenuImages.Draw(pDC, CMenuImages::IdClose, rectSpin, clrImage);
	}
	else
	{
		spinRectUP.bottom -= ((rectSpin.Width() / 2) + 1);
		spinRectDown.top += ((rectSpin.Width() / 2) + 1);

		pDC->FillSolidRect(spinRectUP, clrBackgroundUp);
		pDC->FillSolidRect(spinRectDown, clrBackgroundDown);


		NVS_DrawRectBorder(pDC, spinRectUP, clrFrameUP, 1, 1, 1, 1);
		NVS_DrawRectBorder(pDC, spinRectDown, clrFrameDown, 1, 1, 1, 1);

		currentMenuImages.Draw(pDC, CMenuImages::IdArrowUp, spinRectUP, clrImage);
		currentMenuImages.Draw(pDC, CMenuImages::IdArrowDown, spinRectDown, clrImage);

		rgnMainUP.DeleteObject();
		rgnMainDown.DeleteObject();
		rgnExcludeUP.DeleteObject();
		rgnExcludeDown.DeleteObject();
	}
}

BOOL CNPUP_VisualManager::GetToolTipInfo(CMFCToolTipInfo & params, UINT nType)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;

	params.m_nGradientAngle = TRUE;
	params.m_clrBorder = clrGlblElem.clr_ToolTipInfoBorder;
	params.m_clrFill = clrGlblElem.clr_ToolTipInfoBackground;
	params.m_clrFillGradient = clrGlblElem.clr_ToolTipInfoGradient;
	params.m_clrText = clrGlblElem.clr_ToolTipInfoText;

	return TRUE;
}

void CNPUP_VisualManager::OnFillMenuImageRect
(
	CDC * pDC,
	CMFCToolBarButton * pButton,
	CRect rect,
	CMFCVisualManager::AFX_BUTTON_STATE state
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	CMenuImages currentMenuImages = m_CMIUserImage;

	CMenuImages::IMAGES_IDS drawImage = (CMenuImages::IMAGES_IDS)pButton->GetImage();

	COLORREF clrBackground = clrGlblElem.clr_CheckBoxBackground;

	if (state == AFX_BUTTON_STATE::ButtonsIsRegular)
	{
		clrBackground = clrGlblElem.clr_CheckBoxBackground;
	}
	else if (state == AFX_BUTTON_STATE::ButtonsIsPressed)
	{
		clrBackground = clrGlblElem.clr_CheckBoxBackground_Pressed;
	}
	else if (state == AFX_BUTTON_STATE::ButtonsIsHighlighted)
	{
		clrBackground = clrGlblElem.clr_CheckBoxBackground_Highlight;
	}
	else
	{
		clrBackground = clrGlblElem.clr_CheckBoxBackground;
	}


	pDC->FillSolidRect(rect, RGB(64, 64, 64));

}
#pragma endregion

#pragma endregion


// ============================================================================
// ЭЛЕМЕНТЫ ЛЕНТОЧНОГО ИНТЕРФЕЙСА / RIBBON UI FUNCTIONS
// ----------------------------------------------------------------------------
// Специализированные методы отрисовки компонентов Ribbon-панелей.
// Управляет визуализацией вкладок, категорий, панелей быстрого доступа и 
// меню Ribbon.
//
// Specialized rendering methods for Ribbon interface components.
// Manages tabs, categories, Quick Access Toolbar, and Ribbon-specific menus.
// ============================================================================
#pragma region Ribbon

#pragma region MainFrameRBBN
void CNPUP_VisualManager::OnDrawRibbonCaption
(
	CDC * pDC,
	CMFCRibbonBar * pBar,
	CRect rectCaption,
	CRect rectText
)
{
	rectCaption.top += 1;

	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	ColorTheme::ParametrsCMenuImage prmDraw = m_currentTheme.m_sParamCMenu;
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	CMenuImages currentMenuImages = m_CMIUserImage;

	COLORREF clrTextDraw = clrGlblElem.clr_TextFrame;

	CRect rectButton;
	CRect rectButtonMax;
	CRect rectButtonMin;
	CRect rectButtonCls;
	CString strText;

	CFrameWnd *parrentFrame;

	parrentFrame = pBar->GetParentFrame();


	if (parrentFrame != NULL)
	{
		parrentFrame->GetWindowText(strText);
	}

	HWND hWnd = AfxGetMainWnd()->GetSafeHwnd();
	DWORD style = ::GetWindowLong(hWnd, GWL_STYLE);

	rectButtonMax.CopyRect(rectCaption);
	rectButtonMin.CopyRect(rectCaption);
	rectButtonCls.CopyRect(rectCaption);

	rectButtonCls.left = rectButtonCls.right - (rectCaption.Height() + 18);
	rectButtonMax.right -= 45;
	rectButtonMax.left = rectButtonCls.right - (rectCaption.Height() + 63);
	rectButtonMin.right -= 90;
	rectButtonMin.left = rectButtonMax.right - (rectCaption.Height() + 63);


	pDC->FillSolidRect(rectCaption, clrGlblElem.clr_BackgroundFrame);

	if (style & WS_MAXIMIZE)
	{
		rectText.top = rectCaption.bottom / 2 - 3;

		COLORREF oldClrTextDraw = pDC->SetTextColor(clrTextDraw);

		pDC->DrawText(strText, rectText, 0);

		pDC->SetTextColor(oldClrTextDraw);

		currentMenuImages.Draw(pDC, CMenuImages::IMAGES_IDS::IdRestore, rectButtonMax, clrImage);
	}
	else
	{
		rectText.top = rectCaption.bottom / 4;

		COLORREF oldClrTextDraw = pDC->SetTextColor(clrTextDraw);

		pDC->DrawText(strText, rectText, 0);

		pDC->SetTextColor(oldClrTextDraw);

		currentMenuImages.Draw(pDC, CMenuImages::IMAGES_IDS::IdMaximize, rectButtonMax, clrImage);
	}

	currentMenuImages.Draw(pDC, CMenuImages::IMAGES_IDS::IdClose, rectButtonCls, clrImage);
	currentMenuImages.Draw(pDC, CMenuImages::IMAGES_IDS::IdMinimize, rectButtonMin, clrImage);

	NVS_DrawAlphaChanel(pDC, 255);

}

COLORREF CNPUP_VisualManager::OnDrawRibbonCategoryCaption
(
	CDC * pDC,
	CMFCRibbonContextCaption * pContextCaption
)
{
	ColorTheme::ColorRibbonBackground clrFrameRibbon = m_currentTheme.m_ColorRibbonBackground;
	ColorTheme::ParametrsDrawRibbon prmDrawRibbon = m_currentTheme.m_sParamRibbon;

	CRect rectContext = pContextCaption->GetRect();

	COLORREF clrContext = RGB(0, 0, 0);
	COLORREF clrContextFill = RGB(0, 0, 0);
	COLORREF clrBorder = RGB(0, 0, 0);
	COLORREF clrReturn = RGB(255, 255, 0); //неизвестно

	AFX_RibbonCategoryColor colorContextTab = pContextCaption->GetColor();

	BOOL bIsHighlighted = pContextCaption->IsHighlighted();

	HWND hWnd = AfxGetMainWnd()->GetSafeHwnd();
	DWORD style = ::GetWindowLong(hWnd, GWL_STYLE);

	if (style & WS_MAXIMIZE)
	{
		rectContext.top += 3;
	}

	rectContext.top += 1;

	if (colorContextTab == AFX_CategoryColor_None)
	{
		if (bIsHighlighted)
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Highlight;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight;
		}
		else
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder;
		}
	}
	else if (colorContextTab == AFX_CategoryColor_Red)
	{
		if (bIsHighlighted)
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Red;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Red;
		}
		else
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Red;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Red;
		}
	}
	else if (colorContextTab == AFX_CategoryColor_Orange)
	{
		if (bIsHighlighted)
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Orange;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Orange;
		}
		else
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Orange;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Orange;
		}
	}
	else if (colorContextTab == AFX_CategoryColor_Yellow)
	{
		if (bIsHighlighted)
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Yellow;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Yellow;
		}
		else
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Yellow;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Yellow;
		}
	}
	else if (colorContextTab == AFX_CategoryColor_Green)
	{
		if (bIsHighlighted)
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Green;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Green;
		}
		else
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Green;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Green;
		}
	}
	else if (colorContextTab == AFX_CategoryColor_Blue)
	{
		if (bIsHighlighted)
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Blue;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Blue;
		}
		else
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Blue;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Blue;
		}
	}
	else if (colorContextTab == AFX_CategoryColor_Indigo)
	{
		if (bIsHighlighted)
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Indigo;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Indigo;
		}
		else
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Indigo;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Indigo;
		}
	}
	else if (colorContextTab == AFX_CategoryColor_Violet)
	{
		if (bIsHighlighted)
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Highlight_Violet;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight_Violet;
		}
		else
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Violet;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Violet;
		}
	}
	else
	{
		if (bIsHighlighted)
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption_Highlight;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder_Highlight;
		}
		else
		{
			clrContext = clrFrameRibbon.clr_RibbonCategoryCaption;
			clrBorder = clrFrameRibbon.clr_RibbonCategoryCaptionBorder;
		}
	}

	pDC->FillSolidRect(rectContext, clrContext);

	if (prmDrawRibbon.m_iSizeBorder_ContextBottom > 0
		|| prmDrawRibbon.m_iSizeBorder_ContextLeft > 0
		|| prmDrawRibbon.m_iSizeBorder_ContextRight > 0
		|| prmDrawRibbon.m_iSizeBorder_ContextTop > 0)
	{
		NVS_DrawRectBorder
		(
			pDC,
			rectContext,
			clrBorder,
			prmDrawRibbon.m_iSizeBorder_ContextTop,
			prmDrawRibbon.m_iSizeBorder_ContextBottom,
			prmDrawRibbon.m_iSizeBorder_ContextLeft,
			prmDrawRibbon.m_iSizeBorder_ContextRight
		);
	}

	NVS_DrawAlphaChanel(pDC, 255);

	return clrReturn;
}

COLORREF CNPUP_VisualManager::OnDrawRibbonStatusBarPane
(
	CDC * pDC,
	CMFCRibbonStatusBar * pBar,
	CMFCRibbonStatusBarPane * pPane
)
{
	ColorTheme::ColorStatusBarPane clrStatusBarPane = m_currentTheme.m_ColorStatusBarPane;

	CRect rectPane = pPane->GetRect();

	COLORREF clrPaneBackground = clrStatusBarPane.clr_StatusBarPane_Background;
	COLORREF clrPaneText = clrStatusBarPane.clr_StatusBarPane_Text;

	BOOL bIsHighlight = pPane->IsHighlighted();
	BOOL bIsDisable = pPane->IsDisabled();

	if (bIsDisable)
	{
		clrPaneBackground = clrStatusBarPane.clr_StatusBarPane_Background_Disable;
		clrPaneText = clrStatusBarPane.clr_StatusBarPane_Text_Disable;
	}
	else
	{
		if (bIsHighlight)
		{
			clrPaneBackground = clrStatusBarPane.clr_StatusBarPane_Background_Highlight;
			clrPaneText = clrStatusBarPane.clr_StatusBarPane_Text_Highlight;
		}
		else
		{
			clrPaneBackground = clrStatusBarPane.clr_StatusBarPane_Background;
			clrPaneText = clrStatusBarPane.clr_StatusBarPane_Text;
		}
	}

	pDC->FillSolidRect(rectPane, clrPaneBackground);

	return clrPaneText;
}

#pragma endregion

#pragma region TabsRBBN
COLORREF CNPUP_VisualManager::OnDrawRibbonTabsFrame
(
	CDC * pDC,
	CMFCRibbonBar * pWndRibbonBar,
	CRect rectTab
)
{
	ColorTheme::ColorRibbonTabs clrRibbonTabs = m_currentTheme.m_ColorRibbonTabs;
	ColorTheme::ColorRibbonBackground clrFrameRibbon = m_currentTheme.m_ColorRibbonBackground;
	ColorTheme::ParametrsDrawRibbon prmDrawRibbon = m_currentTheme.m_sParamRibbon;

	COLORREF colorBackgroundRibbon = clrFrameRibbon.clr_RibbonBackground;
	COLORREF colorReturn = RGB(255, 0, 0);

	HWND hWnd = AfxGetMainWnd()->GetSafeHwnd();
	DWORD style = ::GetWindowLong(hWnd, GWL_STYLE);

	if (style & WS_MAXIMIZE)
	{
		rectTab.bottom -= 4;
	}

	pDC->FillSolidRect(rectTab, colorBackgroundRibbon);

	int countCategory = pWndRibbonBar->GetCategoryCount();
	int maxIndexCategory = countCategory - 0;

	for (int countCicle = 0; countCicle < maxIndexCategory; countCicle++)
	{
		CMFCRibbonCategory *categoryRibbon = pWndRibbonBar->GetCategory(countCicle);
		CMFCRibbonTab *tabRibbon = categoryRibbon->GetTab();

		if (categoryRibbon && tabRibbon)
		{
			AFX_RibbonCategoryColor colorTab = categoryRibbon->GetTabColor();
			CRect rectCicleTab = tabRibbon->GetRect();

			COLORREF clrTabBackground = clrRibbonTabs.clr_RibbonTabs;
			COLORREF clrTabBorder = clrRibbonTabs.clr_RibbonTabs;

			if (colorTab == AFX_CategoryColor_None)
			{
				clrTabBackground = clrRibbonTabs.clr_RibbonTabs;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder;
			}
			else if (colorTab == AFX_CategoryColor_Red)
			{
				clrTabBackground = clrRibbonTabs.clr_RibbonTabs_Red;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Red;
			}
			else if (colorTab == AFX_CategoryColor_Orange)
			{
				clrTabBackground = clrRibbonTabs.clr_RibbonTabs_Orange;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Orange;
			}
			else if (colorTab == AFX_CategoryColor_Yellow)
			{
				clrTabBackground = clrRibbonTabs.clr_RibbonTabs_Yellow;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Yellow;
			}
			else if (colorTab == AFX_CategoryColor_Green)
			{
				clrTabBackground = clrRibbonTabs.clr_RibbonTabs_Green;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Green;
			}
			else if (colorTab == AFX_CategoryColor_Blue)
			{
				clrTabBackground = clrRibbonTabs.clr_RibbonTabs_Blue;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Blue;
			}
			else if (colorTab == AFX_CategoryColor_Indigo)
			{
				clrTabBackground = clrRibbonTabs.clr_RibbonTabs_Indigo;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Indigo;
			}
			else if (colorTab == AFX_CategoryColor_Violet)
			{
				clrTabBackground = clrRibbonTabs.clr_RibbonTabs_Violet;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Violet;
			}
			else
			{
				clrTabBackground = clrRibbonTabs.clr_RibbonTabs;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder;
			}

			pDC->FillSolidRect(rectCicleTab, clrTabBackground);

			if (prmDrawRibbon.m_iSizeBorder_TabsTop > 0
				|| prmDrawRibbon.m_iSizeBorder_TabsBottom > 0
				|| prmDrawRibbon.m_iSizeBorder_TabsRight > 0
				|| prmDrawRibbon.m_iSizeBorder_TabsLeft > 0)
			{
				if (prmDrawRibbon.m_bParamRibbon_DrawTabsBackground_BorderAllTabsCaption)
				{
					NVS_DrawRectBorder(pDC, rectCicleTab, clrTabBorder,
						prmDrawRibbon.m_iSizeBorder_TabsTop,
						prmDrawRibbon.m_iSizeBorder_TabsBottom,
						prmDrawRibbon.m_iSizeBorder_TabsLeft,
						prmDrawRibbon.m_iSizeBorder_TabsRight);
				}
				else
				{
					if (prmDrawRibbon.m_bParamRibbon_DrawTabsBackground_BorderNoneTabs
						|| prmDrawRibbon.m_bParamRibbon_DrawTabsBackground_BorderColorTabs)
					{
						if (prmDrawRibbon.m_bParamRibbon_DrawTabsBackground_BorderNoneTabs)
						{
							if (colorTab == AFX_CategoryColor_None)
							{
								NVS_DrawRectBorder(pDC, rectCicleTab, clrTabBorder,
									prmDrawRibbon.m_iSizeBorder_TabsTop,
									prmDrawRibbon.m_iSizeBorder_TabsBottom,
									prmDrawRibbon.m_iSizeBorder_TabsLeft,
									prmDrawRibbon.m_iSizeBorder_TabsRight);
							}
						}

						if (prmDrawRibbon.m_bParamRibbon_DrawTabsBackground_BorderColorTabs)
						{
							if (colorTab != AFX_CategoryColor_None)
							{
								NVS_DrawRectBorder(pDC, rectCicleTab, clrTabBorder,
									prmDrawRibbon.m_iSizeBorder_TabsTop,
									prmDrawRibbon.m_iSizeBorder_TabsBottom,
									prmDrawRibbon.m_iSizeBorder_TabsLeft,
									prmDrawRibbon.m_iSizeBorder_TabsRight);
							}
						}
					}
				}
			}
		}
	}

	return colorReturn;
}

void CNPUP_VisualManager::OnDrawRibbonCategory
(
	CDC * pDC,
	CMFCRibbonCategory * pCategory,
	CRect rectCategory
)
{
	TRACE(_T("Real Class Name: %hs\n"), typeid(*pCategory).name());

	ColorTheme::ColorRibbonCategory clrRibbonCategory = m_currentTheme.m_ColorRibbonCategory;
	ColorTheme::ParametrsDrawRibbon prmDrawRibbon = m_currentTheme.m_sParamRibbon;

	COLORREF colorBackground = RGB(0, 0, 0);
	COLORREF colorFrame = RGB(0, 0, 0);

	AFX_RibbonCategoryColor colorTab = pCategory->GetTabColor();
	CMFCRibbonBar *ribbonParrent = NULL;

	ribbonParrent = pCategory->GetParentRibbonBar();

	if (ribbonParrent == NULL)
	{
		return;
	}

	if (colorTab == AFX_CategoryColor_None)
	{
		colorBackground = clrRibbonCategory.clr_RibbonCategoryBackground;
		colorFrame = clrRibbonCategory.clr_RibbonCategoryBorder;
	}
	else if (colorTab == AFX_CategoryColor_Red)
	{
		colorBackground = clrRibbonCategory.clr_RibbonCategoryBackground_Red;
		colorFrame = clrRibbonCategory.clr_RibbonCategoryBorder_Red;
	}
	else if (colorTab == AFX_CategoryColor_Orange)
	{
		colorBackground = clrRibbonCategory.clr_RibbonCategoryBackground_Orange;
		colorFrame = clrRibbonCategory.clr_RibbonCategoryBorder_Orange;
	}
	else if (colorTab == AFX_CategoryColor_Yellow)
	{
		colorBackground = clrRibbonCategory.clr_RibbonCategoryBackground_Yellow;
		colorFrame = clrRibbonCategory.clr_RibbonCategoryBorder_Yellow;
	}
	else if (colorTab == AFX_CategoryColor_Green)
	{
		colorBackground = clrRibbonCategory.clr_RibbonCategoryBackground_Green;
		colorFrame = clrRibbonCategory.clr_RibbonCategoryBorder_Green;
	}
	else if (colorTab == AFX_CategoryColor_Blue)
	{
		colorBackground = clrRibbonCategory.clr_RibbonCategoryBackground_Blue;
		colorFrame = clrRibbonCategory.clr_RibbonCategoryBorder_Blue;
	}
	else if (colorTab == AFX_CategoryColor_Indigo)
	{
		colorBackground = clrRibbonCategory.clr_RibbonCategoryBackground_Indigo;
		colorFrame = clrRibbonCategory.clr_RibbonCategoryBorder_Indigo;
	}
	else if (colorTab == AFX_CategoryColor_Violet)
	{
		colorBackground = clrRibbonCategory.clr_RibbonCategoryBackground_Violet;
		colorFrame = clrRibbonCategory.clr_RibbonCategoryBorder_Violet;
	}
	else
	{
		colorBackground = clrRibbonCategory.clr_RibbonCategoryBackground;
		colorFrame = clrRibbonCategory.clr_RibbonCategoryBorder;
	}

	if (ribbonParrent->GetHideFlags())
	{
		CBitmap *bmp = new CBitmap();
		bmp->CreateCompatibleBitmap(pDC, rectCategory.Width(), rectCategory.Height());
		CBitmap* pOldBmp = pDC->SelectObject(bmp);
		pOldBmp->DeleteObject();
	}

	pDC->FillSolidRect(rectCategory, colorBackground);

	if (prmDrawRibbon.m_iSizeBorder_CategoryTop > 0
		|| prmDrawRibbon.m_iSizeBorder_CategoryBottom > 0
		|| prmDrawRibbon.m_iSizeBorder_CategoryRight > 0
		|| prmDrawRibbon.m_iSizeBorder_CategoryLeft > 0)
	{
		NVS_DrawRectBorder(
			pDC, rectCategory, colorFrame,
			prmDrawRibbon.m_iSizeBorder_CategoryTop,
			prmDrawRibbon.m_iSizeBorder_CategoryBottom,
			prmDrawRibbon.m_iSizeBorder_CategoryLeft,
			prmDrawRibbon.m_iSizeBorder_CategoryRight);
	}

}

COLORREF CNPUP_VisualManager::OnDrawRibbonCategoryTab
(
	CDC * pDC,
	CMFCRibbonTab * pTab,
	BOOL bIsActive
)
{
	ColorTheme::ColorRibbonTabs clrRibbonTabs = m_currentTheme.m_ColorRibbonTabs;
	ColorTheme::ParametrsDrawRibbon prmDrawRibbon = m_currentTheme.m_sParamRibbon;

	COLORREF clrTab = RGB(0, 0, 0);
	COLORREF clrTabBorder = RGB(0, 0, 0);
	COLORREF clrTextTab = RGB(0, 0, 0);
	CRect rectTab;

	AFX_RibbonCategoryColor colorCategoryTab = AFX_CategoryColor_None;
	CMFCRibbonCategory *categoryParrent = NULL;
	CMFCRibbonBar *ribbonParrent = NULL;

	BOOL highlightedTab = pTab->IsHighlighted();
	BOOL pressedTab = pTab->IsPressed();
	BOOL activeTab = bIsActive;

	BOOL hideRibbonTab = FALSE;

	rectTab = pTab->GetRect();

	categoryParrent = pTab->GetParentCategory();
	if (categoryParrent == NULL)
	{
		return clrTextTab;
	}

	ribbonParrent = categoryParrent->GetParentRibbonBar();
	if (ribbonParrent == NULL)
	{
		return clrTextTab;
	}

	colorCategoryTab = categoryParrent->GetTabColor();

	//---------GhostTab---------
	if (pressedTab == FALSE
		&& highlightedTab == FALSE
		&& activeTab == FALSE)
	{
		rectTab.left = 0;
		rectTab.right = 0;
		rectTab.bottom = 0;
		rectTab.top = 0;
	}
	//--------------------------

	if (ribbonParrent->GetHideFlags())
	{
		if (ribbonParrent->GetDroppedDown() == NULL)
		{
			activeTab = FALSE;
		}
		else if (ribbonParrent->GetDroppedDown() != NULL)
		{
			activeTab = TRUE;
		}

		hideRibbonTab = TRUE;
	}
	else
	{
		hideRibbonTab = FALSE;
	}


	if (colorCategoryTab == AFX_CategoryColor_None)
	{
		if (highlightedTab)
		{
			clrTab = clrRibbonTabs.clr_RibbonTabs_Highlight;
			clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Highlight;
			clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Highlight;
		}
		else
		{
			if (activeTab)
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Active;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Active;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Active;
			}
			else
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder;
			}
		}
	}
	else if (colorCategoryTab == AFX_CategoryColor_Red)
	{
		if (highlightedTab)
		{
			clrTab = clrRibbonTabs.clr_RibbonTabs_Highlight_Red;
			clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Red;
			clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Red;
		}
		else
		{
			if (activeTab)
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Active_Red;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Active_Red;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Active_Red;
			}
			else
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Red;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Red;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Red;
			}
		}
	}
	else if (colorCategoryTab == AFX_CategoryColor_Orange)
	{
		if (highlightedTab)
		{
			clrTab = clrRibbonTabs.clr_RibbonTabs_Highlight_Orange;
			clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Orange;
			clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Orange;
		}
		else
		{
			if (activeTab)
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Active_Orange;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Active_Orange;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Active_Orange;
			}
			else
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Orange;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Orange;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Orange;
			}
		}
	}
	else if (colorCategoryTab == AFX_CategoryColor_Yellow)
	{
		if (highlightedTab)
		{
			clrTab = clrRibbonTabs.clr_RibbonTabs_Highlight_Yellow;
			clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Yellow;
			clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Yellow;
		}
		else
		{
			if (activeTab)
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Active_Yellow;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Active_Yellow;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Active_Yellow;
			}
			else
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Yellow;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Yellow;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Yellow;
			}
		}
	}
	else if (colorCategoryTab == AFX_CategoryColor_Green)
	{
		if (highlightedTab)
		{
			clrTab = clrRibbonTabs.clr_RibbonTabs_Highlight_Green;
			clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Green;
			clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Green;
		}
		else
		{
			if (activeTab)
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Active_Green;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Active_Green;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Active_Green;
			}
			else
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Green;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Green;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Green;
			}
		}
	}
	else if (colorCategoryTab == AFX_CategoryColor_Blue)
	{
		if (highlightedTab)
		{
			clrTab = clrRibbonTabs.clr_RibbonTabs_Highlight_Blue;
			clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Blue;
			clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Blue;
		}
		else
		{
			if (activeTab)
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Active_Blue;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Active_Blue;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Active_Blue;
			}
			else
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Blue;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Blue;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Blue;
			}
		}
	}
	else if (colorCategoryTab == AFX_CategoryColor_Indigo)
	{
		if (highlightedTab)
		{
			clrTab = clrRibbonTabs.clr_RibbonTabs_Highlight_Indigo;
			clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Indigo;
			clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Indigo;
		}
		else
		{
			if (activeTab)
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Active_Indigo;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Active_Indigo;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Active_Indigo;
			}
			else
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Indigo;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Indigo;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Indigo;
			}
		}
	}
	else if (colorCategoryTab == AFX_CategoryColor_Violet)
	{
		if (highlightedTab)
		{
			clrTab = clrRibbonTabs.clr_RibbonTabs_Highlight_Violet;
			clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Highlight_Violet;
			clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Highlight_Violet;
		}
		else
		{
			if (activeTab)
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Active_Violet;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Active_Violet;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Active_Violet;
			}
			else
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Violet;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Violet;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Violet;
			}
		}
	}
	else
	{
		if (highlightedTab)
		{
			clrTab = clrRibbonTabs.clr_RibbonTabs_Highlight;
			clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Highlight;
			clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Highlight;
		}
		else
		{
			if (activeTab)
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs_Active;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text_Active;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder_Active;
			}
			else
			{
				clrTab = clrRibbonTabs.clr_RibbonTabs;
				clrTextTab = clrRibbonTabs.clr_RibbonTabs_Text;
				clrTabBorder = clrRibbonTabs.clr_RibbonTabsBorder;
			}
		}
	}

	pDC->FillSolidRect(rectTab, clrTab);

	if (prmDrawRibbon.m_iSizeBorder_TabsTop > 0
		|| prmDrawRibbon.m_iSizeBorder_TabsBottom > 0
		|| prmDrawRibbon.m_iSizeBorder_TabsRight > 0
		|| prmDrawRibbon.m_iSizeBorder_TabsLeft > 0)
	{
		if (prmDrawRibbon.m_bParamRibbon_DrawTabsBackground_BorderAllTabsCaption)
		{
			NVS_DrawRectBorder(pDC, rectTab, clrTabBorder,
				prmDrawRibbon.m_iSizeBorder_TabsTop,
				prmDrawRibbon.m_iSizeBorder_TabsBottom,
				prmDrawRibbon.m_iSizeBorder_TabsLeft,
				prmDrawRibbon.m_iSizeBorder_TabsRight);
		}
		else
		{
			if (prmDrawRibbon.m_bParamRibbon_DrawTabsBackground_BorderNoneTabs
				|| prmDrawRibbon.m_bParamRibbon_DrawTabsBackground_BorderColorTabs)
			{
				if (prmDrawRibbon.m_bParamRibbon_DrawTabsBackground_BorderNoneTabs)
				{
					if (colorCategoryTab == AFX_CategoryColor_None)
					{
						NVS_DrawRectBorder(pDC, rectTab, clrTabBorder,
							prmDrawRibbon.m_iSizeBorder_TabsTop,
							prmDrawRibbon.m_iSizeBorder_TabsBottom,
							prmDrawRibbon.m_iSizeBorder_TabsLeft,
							prmDrawRibbon.m_iSizeBorder_TabsRight);
					}
				}

				if (prmDrawRibbon.m_bParamRibbon_DrawTabsBackground_BorderColorTabs)
				{
					if (colorCategoryTab != AFX_CategoryColor_None)
					{
						NVS_DrawRectBorder(pDC, rectTab, clrTabBorder,
							prmDrawRibbon.m_iSizeBorder_TabsTop,
							prmDrawRibbon.m_iSizeBorder_TabsBottom,
							prmDrawRibbon.m_iSizeBorder_TabsLeft,
							prmDrawRibbon.m_iSizeBorder_TabsRight);
					}
				}
			}
		}
	}


	return clrTextTab;
}

COLORREF CNPUP_VisualManager::OnDrawRibbonPanel
(
	CDC * pDC,
	CMFCRibbonPanel * pPanel,
	CRect rectPanel,
	CRect rectCaption
)
{
	ColorTheme::ColorRibbonPanel clrRibbonPanel = m_currentTheme.m_ColorRibbonPanel;
	ColorTheme::ColorMainMenuRibbon clrMainMenuRibbon = m_currentTheme.m_ColorMainMenuRibbon;
	ColorTheme::ParametrsDrawRibbon prmDrawRbn = m_currentTheme.m_sParamRibbon;

	COLORREF colorBackgroundPanel = RGB(0, 0, 0);
	COLORREF colorBorder = RGB(0, 0, 0);
	COLORREF colorText = RGB(0, 0, 0);

	CMFCRibbonCategory *categoryParrent = NULL;
	AFX_RibbonCategoryColor colorCategoryTab = AFX_CategoryColor_None;

	BOOL bIsHighlight = pPanel->IsHighlighted();
	BOOL bIsMenuMode = pPanel->IsMenuMode();

	categoryParrent = pPanel->GetParentCategory();
	if (categoryParrent == NULL)
	{
		return colorText;
	}

	colorCategoryTab = categoryParrent->GetTabColor();

	if (bIsMenuMode)
	{
		colorBackgroundPanel = clrMainMenuRibbon.clr_RibbonMainMenuMainPanelBackground;
		colorBorder = clrMainMenuRibbon.clr_RibbonMainMenuMainPanelBackground;
		colorText = clrMainMenuRibbon.clr_RibbonMainMenuButtonText;
	}
	else
	{
		if (colorCategoryTab == AFX_CategoryColor_None)
		{
			if (bIsHighlight)
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight;
			}
			else
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText;
			}
		}
		else if (colorCategoryTab == AFX_CategoryColor_Red)
		{
			if (bIsHighlight)
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Red;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Red;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Red;
			}
			else
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Red;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Red;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Red;
			}
		}
		else if (colorCategoryTab == AFX_CategoryColor_Orange)
		{
			if (bIsHighlight)
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Orange;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Orange;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Orange;
			}
			else
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Orange;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Orange;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Orange;
			}
		}
		else if (colorCategoryTab == AFX_CategoryColor_Yellow)
		{
			if (bIsHighlight)
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Yellow;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Yellow;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Yellow;
			}
			else
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Yellow;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Yellow;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Yellow;
			}
		}
		else if (colorCategoryTab == AFX_CategoryColor_Green)
		{
			if (bIsHighlight)
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Green;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Green;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Green;
			}
			else
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Green;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Green;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Green;
			}
		}
		else if (colorCategoryTab == AFX_CategoryColor_Blue)
		{
			if (bIsHighlight)
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Blue;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Blue;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Blue;
			}
			else
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Blue;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Blue;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Blue;
			}
		}
		else if (colorCategoryTab == AFX_CategoryColor_Indigo)
		{
			if (bIsHighlight)
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Indigo;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Indigo;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Indigo;
			}
			else
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Indigo;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Indigo;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Indigo;
			}
		}
		else if (colorCategoryTab == AFX_CategoryColor_Violet)
		{
			if (bIsHighlight)
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight_Violet;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight_Violet;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight_Violet;
			}
			else
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Violet;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Violet;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Violet;
			}
		}
		else
		{
			if (bIsHighlight)
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground_Highlight;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder_Highlight;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Highlight;
			}
			else
			{
				colorBackgroundPanel = clrRibbonPanel.clr_RibbonPanelTabBackground;
				colorBorder = clrRibbonPanel.clr_RibbonPanelTabBorder;
				colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText;
			}
		}
	}

	pDC->FillSolidRect(rectPanel, colorBackgroundPanel);
	// pDC->FillSolidRect(rectCaption, RGB(255, 0, 0)); // ПЕРЕНЕСЕНО В OnDrawRibbonPanelCaption

	if (prmDrawRbn.m_iSizeBorder_PanelTop > 0
		|| prmDrawRbn.m_iSizeBorder_PanelBottom > 0
		|| prmDrawRbn.m_iSizeBorder_PanelRight > 0
		|| prmDrawRbn.m_iSizeBorder_PanelLeft > 0)
	{
		NVS_DrawRectBorder
		(pDC, rectPanel, colorBorder,
			prmDrawRbn.m_iSizeBorder_PanelTop,
			prmDrawRbn.m_iSizeBorder_PanelBottom,
			prmDrawRbn.m_iSizeBorder_PanelLeft,
			prmDrawRbn.m_iSizeBorder_PanelRight
		);
	}


	return colorText;
}

void CNPUP_VisualManager::OnDrawRibbonPanelCaption
(
	CDC * pDC,
	CMFCRibbonPanel * pPanel,
	CRect rectCaption
)
{
	ColorTheme::ColorRibbonPanel clrRibbonPanel = m_currentTheme.m_ColorRibbonPanel;
	ColorTheme::ParametrsDrawRibbon prmDraw = m_currentTheme.m_sParamRibbon;

	COLORREF colorBackgroundCaptionPanel = RGB(0, 0, 0);
	COLORREF colorBorderCaptionPanel = RGB(0, 0, 0);
	COLORREF colorText = RGB(0, 0, 0);
	COLORREF oldColor;

	CMFCRibbonCategory *categoryParrent = NULL;
	AFX_RibbonCategoryColor colorCategoryTab = AFX_CategoryColor_None;

	BOOL bIsHighlighted = pPanel->IsHighlighted(); // потом
	BOOL bIsMenuMode = pPanel->IsMenuMode();

	categoryParrent = pPanel->GetParentCategory();
	if (categoryParrent == NULL)
	{
		return;
	}

	colorCategoryTab = categoryParrent->GetTabColor();

	if (bIsMenuMode)
	{
		colorBackgroundCaptionPanel = RGB(0, 0, 255);
		colorBorderCaptionPanel = RGB(0, 255, 0);
		colorText = RGB(255, 0, 0);
	}
	else
	{
		if (colorCategoryTab == AFX_CategoryColor_None)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground;
			colorBorderCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBorder;
			colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Red)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Red;
			colorBorderCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Red;
			colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Red;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Orange)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Orange;
			colorBorderCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Orange;
			colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Orange;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Yellow)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Yellow;
			colorBorderCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Yellow;
			colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Yellow;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Green)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Green;
			colorBorderCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Green;
			colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Green;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Blue)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Blue;
			colorBorderCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Blue;
			colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Blue;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Indigo)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Indigo;
			colorBorderCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Indigo;
			colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Indigo;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Violet)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Violet;
			colorBorderCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBorder_Violet;
			colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText_Violet;
		}
		else
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground;
			colorBorderCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBorder;
			colorText = clrRibbonPanel.clr_RibbonPanelTabLabelText;
		}
	}

	pDC->FillSolidRect(rectCaption, colorBackgroundCaptionPanel); //еще один фон подписей категорий внутри вкладки ленты, но он перекрывает и текст

	oldColor = pDC->SetTextColor(colorText);

	pDC->DrawText(pPanel->GetName(), rectCaption, TRANSPARENT);

	pDC->SetTextColor(oldColor);

	if (prmDraw.m_iSizeBorder_PanelLabelTop > 0
		|| prmDraw.m_iSizeBorder_PanelLabelBottom > 0
		|| prmDraw.m_iSizeBorder_PanelLabelRight > 0
		|| prmDraw.m_iSizeBorder_PanelLabelLeft > 0)
	{
		NVS_DrawRectBorder
		(
			pDC, rectCaption, colorBorderCaptionPanel,
			prmDraw.m_iSizeBorder_PanelLabelTop,
			prmDraw.m_iSizeBorder_PanelLabelBottom,
			prmDraw.m_iSizeBorder_PanelLabelLeft,
			prmDraw.m_iSizeBorder_PanelLabelRight
		);
	}

}

COLORREF CNPUP_VisualManager::OnFillRibbonButton
(
	CDC * pDC,
	CMFCRibbonButton * pButton
)
{
	ColorTheme::ColorRibbonButton clrRibbonButton = m_currentTheme.m_ColorRibbonButton;
	ColorTheme::ColorMainMenuRibbon clrMainMenuRibbon = m_currentTheme.m_ColorMainMenuRibbon;
	ColorTheme::ColorRibbonSystemAndQuickButton clrSysQuickButton = m_currentTheme.m_ColorRibbonSystemAndQuickButton;

	CRect rectButton;
	COLORREF colorButton = RGB(0, 0, 255);
	COLORREF colorTextButton = RGB(0, 255, 0);
	AFX_RibbonCategoryColor colorCategory = AFX_CategoryColor_None;

	CMFCRibbonCategory *categoryParrent = NULL;
	CMFCRibbonBar *ribbonParrent = NULL;
	CMFCRibbonPanel *panelParrent = NULL;

	BOOL buttonDisabled = pButton->IsDisabled();
	BOOL buttonHighlighted = pButton->IsHighlighted();
	BOOL buttonPressed = pButton->IsPressed();
	BOOL buttonMenuMode = pButton->IsMenuMode();

	rectButton = pButton->GetRect();

	panelParrent = pButton->GetParentPanel();

	if (pButton->IsMenuMode())
	{
		if (buttonDisabled)
		{
			colorButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground_Disabled;
			colorTextButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonText_Disabled;
		}
		else
		{
			if (buttonPressed)
			{
				colorButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground_Pressed;
				colorTextButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonText_Pressed;
			}
			else
			{
				if (buttonHighlighted)
				{
					colorButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground_Highlight;
					colorTextButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonText_Highlight;
				}
				else
				{
					colorButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground;
					colorTextButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonText;
				}
			}
		}
	}
	else
	{
		if (panelParrent == NULL)
		{
			if (buttonDisabled)
			{
				colorButton = clrSysQuickButton.clr_RibbonSystemButton_Background_Disable;
				colorTextButton = clrSysQuickButton.clr_RibbonSystemButton_Text_Disable;
			}
			else
			{
				if (buttonHighlighted)
				{
					colorButton = clrSysQuickButton.clr_RibbonSystemButton_Background_Highlight;
					colorTextButton = clrSysQuickButton.clr_RibbonSystemButton_Text_Highlight;
				}
				else
				{
					if (buttonPressed)
					{
						colorButton = clrSysQuickButton.clr_RibbonSystemButton_Background_Pressed;
						colorTextButton = clrSysQuickButton.clr_RibbonSystemButton_Text_Pressed;
					}
					else
					{
						colorButton = clrSysQuickButton.clr_RibbonSystemButton_Background;
						colorTextButton = clrSysQuickButton.clr_RibbonSystemButton_Text;
					}
				}
			}
		}
		else
		{
			categoryParrent = panelParrent->GetParentCategory();
			if (categoryParrent == NULL)
			{
				if (buttonDisabled)
				{
					colorButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground_Disabled;
					colorTextButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonText_Disabled;
				}
				else
				{
					if (buttonPressed)
					{
						colorButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground_Pressed;
						colorTextButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonText_Pressed;
					}
					else
					{
						if (buttonHighlighted)
						{
							colorButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground_Highlight;
							colorTextButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonText_Highlight;
						}
						else
						{
							colorButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground;
							colorTextButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonText;
						}
					}
				}
			}
			else
			{
				colorCategory = categoryParrent->GetTabColor();

				if (colorCategory == AFX_CategoryColor_None)
				{
					if (buttonDisabled)
					{
						colorButton = clrRibbonButton.clr_RibbonButtonBackground_Disable;
						colorTextButton = clrRibbonButton.clr_RibbonButtonText_Disable;
					}
					else
					{
						if (buttonHighlighted)
						{
							colorButton = clrRibbonButton.clr_RibbonButtonBackground_Highlight;
							colorTextButton = clrRibbonButton.clr_RibbonButtonText_Highlight;
						}
						else
						{
							if (buttonPressed)
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Pressed;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Pressed;
							}
							else
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText;
							}
						}
					}
				}
				else if (colorCategory == AFX_CategoryColor_Red)
				{
					if (buttonDisabled)
					{
						colorButton = clrRibbonButton.clr_RibbonButtonBackground_Disable_Red;
						colorTextButton = clrRibbonButton.clr_RibbonButtonText_Disable_Red;
					}
					else
					{
						if (buttonHighlighted)
						{
							colorButton = clrRibbonButton.clr_RibbonButtonBackground_Highlight_Red;
							colorTextButton = clrRibbonButton.clr_RibbonButtonText_Highlight_Red;
						}
						else
						{
							if (buttonPressed)
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Pressed_Red;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Pressed_Red;
							}
							else
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Red;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Red;
							}
						}
					}
				}
				else if (colorCategory == AFX_CategoryColor_Orange)
				{
					if (buttonDisabled)
					{
						colorButton = clrRibbonButton.clr_RibbonButtonBackground_Disable_Orange;
						colorTextButton = clrRibbonButton.clr_RibbonButtonText_Disable_Orange;
					}
					else
					{
						if (buttonHighlighted)
						{
							colorButton = clrRibbonButton.clr_RibbonButtonBackground_Highlight_Orange;
							colorTextButton = clrRibbonButton.clr_RibbonButtonText_Highlight_Orange;
						}
						else
						{
							if (buttonPressed)
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Pressed_Orange;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Pressed_Orange;
							}
							else
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Orange;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Orange;
							}
						}
					}
				}
				else if (colorCategory == AFX_CategoryColor_Yellow)
				{
					if (buttonDisabled)
					{
						colorButton = clrRibbonButton.clr_RibbonButtonBackground_Disable_Yellow;
						colorTextButton = clrRibbonButton.clr_RibbonButtonText_Disable_Yellow;
					}
					else
					{
						if (buttonHighlighted)
						{
							colorButton = clrRibbonButton.clr_RibbonButtonBackground_Highlight_Yellow;
							colorTextButton = clrRibbonButton.clr_RibbonButtonText_Highlight_Yellow;
						}
						else
						{
							if (buttonPressed)
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Pressed_Yellow;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Pressed_Yellow;
							}
							else
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Yellow;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Yellow;
							}
						}
					}
				}
				else if (colorCategory == AFX_CategoryColor_Green)
				{
					if (buttonDisabled)
					{
						colorButton = clrRibbonButton.clr_RibbonButtonBackground_Disable_Green;
						colorTextButton = clrRibbonButton.clr_RibbonButtonText_Disable_Green;
					}
					else
					{
						if (buttonHighlighted)
						{
							colorButton = clrRibbonButton.clr_RibbonButtonBackground_Highlight_Green;
							colorTextButton = clrRibbonButton.clr_RibbonButtonText_Highlight_Green;
						}
						else
						{
							if (buttonPressed)
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Pressed_Green;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Pressed_Green;
							}
							else
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Green;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Green;
							}
						}
					}
				}
				else if (colorCategory == AFX_CategoryColor_Blue)
				{
					if (buttonDisabled)
					{
						colorButton = clrRibbonButton.clr_RibbonButtonBackground_Disable_Blue;
						colorTextButton = clrRibbonButton.clr_RibbonButtonText_Disable_Blue;
					}
					else
					{
						if (buttonHighlighted)
						{
							colorButton = clrRibbonButton.clr_RibbonButtonBackground_Highlight_Blue;
							colorTextButton = clrRibbonButton.clr_RibbonButtonText_Highlight_Blue;
						}
						else
						{
							if (buttonPressed)
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Pressed_Blue;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Pressed_Blue;
							}
							else
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Blue;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Blue;
							}
						}
					}
				}
				else if (colorCategory == AFX_CategoryColor_Indigo)
				{
					if (buttonDisabled)
					{
						colorButton = clrRibbonButton.clr_RibbonButtonBackground_Disable_Indigo;
						colorTextButton = clrRibbonButton.clr_RibbonButtonText_Disable_Indigo;
					}
					else
					{
						if (buttonHighlighted)
						{
							colorButton = clrRibbonButton.clr_RibbonButtonBackground_Highlight_Indigo;
							colorTextButton = clrRibbonButton.clr_RibbonButtonText_Highlight_Indigo;
						}
						else
						{
							if (buttonPressed)
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Pressed_Indigo;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Pressed_Indigo;
							}
							else
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Indigo;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Indigo;
							}
						}
					}
				}
				else if (colorCategory == AFX_CategoryColor_Violet)
				{
					if (buttonDisabled)
					{
						colorButton = clrRibbonButton.clr_RibbonButtonBackground_Disable_Violet;
						colorTextButton = clrRibbonButton.clr_RibbonButtonText_Disable_Violet;
					}
					else
					{
						if (buttonHighlighted)
						{
							colorButton = clrRibbonButton.clr_RibbonButtonBackground_Highlight_Violet;
							colorTextButton = clrRibbonButton.clr_RibbonButtonText_Highlight_Violet;
						}
						else
						{
							if (buttonPressed)
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Pressed_Violet;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Pressed_Violet;
							}
							else
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Violet;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Violet;
							}
						}
					}
				}
				else
				{
					if (buttonDisabled)
					{
						colorButton = clrRibbonButton.clr_RibbonButtonBackground_Disable;
						colorTextButton = clrRibbonButton.clr_RibbonButtonText_Disable;
					}
					else
					{
						if (buttonHighlighted)
						{
							colorButton = clrRibbonButton.clr_RibbonButtonBackground_Highlight;
							colorTextButton = clrRibbonButton.clr_RibbonButtonText_Highlight;
						}
						else
						{
							if (buttonPressed)
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground_Pressed;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText_Pressed;
							}
							else
							{
								colorButton = clrRibbonButton.clr_RibbonButtonBackground;
								colorTextButton = clrRibbonButton.clr_RibbonButtonText;
							}
						}
					}
				}
			}
		}
	}

	pDC->FillSolidRect(rectButton, colorButton);


	return colorTextButton;
}

void CNPUP_VisualManager::OnDrawRibbonButtonBorder
(
	CDC * pDC,
	CMFCRibbonButton * pButton
)
{
	ColorTheme::ColorRibbonButton clrRibbonButton = m_currentTheme.m_ColorRibbonButton;
	ColorTheme::ColorMainMenuRibbon clrMainMenuRibbon = m_currentTheme.m_ColorMainMenuRibbon;
	ColorTheme::ParametrsDrawRibbon prmDraw = m_currentTheme.m_sParamRibbon;

	CRect rectButton = pButton->GetRect();
	COLORREF colorBorderButton = RGB(0, 255, 0);
	AFX_RibbonCategoryColor colorCategory = AFX_CategoryColor_None;

	CMFCRibbonCategory *categoryParrent = NULL;
	CMFCRibbonBar *ribbonParrent = NULL;
	CMFCRibbonPanel *panelParrent = NULL;

	BOOL buttonDisabled = pButton->IsDisabled();
	BOOL buttonHighlighted = pButton->IsHighlighted();
	BOOL buttonPressed = pButton->IsPressed();

	panelParrent = pButton->GetParentPanel();
	if (panelParrent == NULL)
	{
		return;
	}

	categoryParrent = panelParrent->GetParentCategory();
	if (categoryParrent == NULL)
	{
		return;
	}

	colorCategory = categoryParrent->GetTabColor();

	if (pButton->IsMenuMode())
	{
		if (buttonDisabled)
		{
			colorBorderButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonBorder_Disabled;
		}
		else
		{
			if (buttonPressed)
			{
				colorBorderButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonBorder_Pressed;
			}
			else
			{
				if (buttonHighlighted)
				{
					colorBorderButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonBorder_Highlight;
				}
				else
				{
					colorBorderButton = clrMainMenuRibbon.clr_RibbonMainMenuButtonBorder;
				}
			}
		}
	}
	else
	{
		if (colorCategory == AFX_CategoryColor_None)
		{
			if (buttonDisabled)
			{
				colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Disable;
			}
			else
			{
				if (buttonHighlighted)
				{
					colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Highlight;
				}
				else
				{
					if (buttonPressed)
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Pressed;
					}
					else
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder;
					}
				}
			}
		}
		else if (colorCategory == AFX_CategoryColor_Red)
		{
			if (buttonDisabled)
			{
				colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Disable_Red;
			}
			else
			{
				if (buttonHighlighted)
				{
					colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Highlight_Red;
				}
				else
				{
					if (buttonPressed)
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Pressed_Red;
					}
					else
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Red;
					}
				}
			}
		}
		else if (colorCategory == AFX_CategoryColor_Orange)
		{
			if (buttonDisabled)
			{
				colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Disable_Orange;
			}
			else
			{
				if (buttonHighlighted)
				{
					colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Highlight_Orange;
				}
				else
				{
					if (buttonPressed)
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Pressed_Orange;
					}
					else
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Orange;
					}
				}
			}
		}
		else if (colorCategory == AFX_CategoryColor_Yellow)
		{
			if (buttonDisabled)
			{
				colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Disable_Yellow;
			}
			else
			{
				if (buttonHighlighted)
				{
					colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Highlight_Yellow;
				}
				else
				{
					if (buttonPressed)
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Pressed_Yellow;
					}
					else
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Yellow;
					}
				}
			}
		}
		else if (colorCategory == AFX_CategoryColor_Green)
		{
			if (buttonDisabled)
			{
				colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Disable_Green;
			}
			else
			{
				if (buttonHighlighted)
				{
					colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Highlight_Green;
				}
				else
				{
					if (buttonPressed)
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Pressed_Green;
					}
					else
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Green;
					}
				}
			}
		}
		else if (colorCategory == AFX_CategoryColor_Blue)
		{
			if (buttonDisabled)
			{
				colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Disable_Blue;
			}
			else
			{
				if (buttonHighlighted)
				{
					colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Highlight_Blue;
				}
				else
				{
					if (buttonPressed)
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Pressed_Blue;
					}
					else
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Blue;
					}
				}
			}
		}
		else if (colorCategory == AFX_CategoryColor_Indigo)
		{
			if (buttonDisabled)
			{
				colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Disable_Indigo;
			}
			else
			{
				if (buttonHighlighted)
				{
					colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Highlight_Indigo;
				}
				else
				{
					if (buttonPressed)
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Pressed_Indigo;
					}
					else
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Indigo;
					}
				}
			}
		}
		else if (colorCategory == AFX_CategoryColor_Violet)
		{
			if (buttonDisabled)
			{
				colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Disable_Violet;
			}
			else
			{
				if (buttonHighlighted)
				{
					colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Highlight_Violet;
				}
				else
				{
					if (buttonPressed)
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Pressed_Violet;
					}
					else
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Violet;
					}
				}
			}
		}
		else
		{
			if (buttonDisabled)
			{
				colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Disable;
			}
			else
			{
				if (buttonHighlighted)
				{
					colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Highlight;
				}
				else
				{
					if (buttonPressed)
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder_Pressed;
					}
					else
					{
						colorBorderButton = clrRibbonButton.clr_RibbonButtonBorder;
					}
				}
			}
		}
	}

	if (prmDraw.m_iSizeBorder_RibbonButtonBottom > 0
		|| prmDraw.m_iSizeBorder_RibbonButtonLeft > 0
		|| prmDraw.m_iSizeBorder_RibbonButtonRight > 0
		|| prmDraw.m_iSizeBorder_RibbonButtonTop > 0)
	{
		NVS_DrawRectBorder
		(pDC, rectButton, colorBorderButton,
			prmDraw.m_iSizeBorder_RibbonButtonTop,
			prmDraw.m_iSizeBorder_RibbonButtonBottom,
			prmDraw.m_iSizeBorder_RibbonButtonLeft,
			prmDraw.m_iSizeBorder_RibbonButtonRight);
	}
}

#pragma endregion

#pragma region MainMenuRBBN
void CNPUP_VisualManager::OnDrawRibbonApplicationButton
(
	CDC * pDC,
	CMFCRibbonButton * pButton
)
{
	ColorTheme::ColorMainMenuRibbon clrMainMenuRibbon = m_currentTheme.m_ColorMainMenuRibbon;

	CRect rectButton = pButton->GetRect();

	COLORREF clrButtonBackground = RGB(255, 0, 0);

	BOOL buttonDisabled = pButton->IsDisabled();
	BOOL buttonHighlighted = pButton->IsHighlighted();
	BOOL buttonPressed = pButton->IsPressed();

	if (buttonDisabled)
	{
		clrButtonBackground = clrMainMenuRibbon.clr_RibbonMainMenuApplicationButtonBackground_Disabled;
	}
	else
	{
		if (buttonHighlighted)
		{
			clrButtonBackground = clrMainMenuRibbon.clr_RibbonMainMenuApplicationButtonBackground_Highlight;
		}
		else
		{
			if (buttonPressed)
			{
				clrButtonBackground = clrMainMenuRibbon.clr_RibbonMainMenuApplicationButtonBackground_Pressed;
			}
			else
			{
				clrButtonBackground = clrMainMenuRibbon.clr_RibbonMainMenuApplicationButtonBackground;
			}
		}
	}

	pDC->FillSolidRect(rectButton, clrButtonBackground);
	NVS_DrawAlphaChanel(pDC, 255);

}

void CNPUP_VisualManager::OnDrawRibbonMainPanelFrame
(
	CDC * pDC,
	CMFCRibbonMainPanel * pPanel,
	CRect rect
)
{
	ColorTheme::ColorMainMenuRibbon clrMainMenuRibbon = m_currentTheme.m_ColorMainMenuRibbon;

	CRect rectPanel;

	COLORREF clrMainPanel = clrMainMenuRibbon.clr_RibbonMainMenuMainPanelBackground;

	rectPanel.CopyRect(rect);
	rectPanel.right += 2;

	pDC->FillSolidRect(rectPanel, clrMainPanel);
}

COLORREF CNPUP_VisualManager::OnFillRibbonMainPanelButton
(
	CDC * pDC,
	CMFCRibbonButton * pButton
)
{
	ColorTheme::ColorMainMenuRibbon clrMainMenuRibbon = m_currentTheme.m_ColorMainMenuRibbon;

	CRect rectButton = pButton->GetRect();

	COLORREF clrButtonText = RGB(255, 0, 0);
	COLORREF clrButtonBackground = RGB(255, 0, 0);

	BOOL buttonDisabled = pButton->IsDisabled();
	BOOL buttonHighlighted = pButton->IsHighlighted();
	BOOL buttonPressed = pButton->IsPressed();

	if (buttonDisabled)
	{
		clrButtonText = clrMainMenuRibbon.clr_RibbonMainMenuButtonText_Disabled;
		clrButtonBackground = clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground_Disabled;
	}
	else
	{
		if (buttonHighlighted)
		{
			clrButtonText = clrMainMenuRibbon.clr_RibbonMainMenuButtonText_Highlight;
			clrButtonBackground = clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground_Highlight;
		}
		else
		{
			if (buttonPressed)
			{
				clrButtonText = clrMainMenuRibbon.clr_RibbonMainMenuButtonText_Pressed;
				clrButtonBackground = clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground_Pressed;
			}
			else
			{
				clrButtonText = clrMainMenuRibbon.clr_RibbonMainMenuButtonText;
				clrButtonBackground = clrMainMenuRibbon.clr_RibbonMainMenuButtonBackground;
			}
		}
	}

	pDC->FillSolidRect(rectButton, clrButtonBackground);

	return clrButtonText;
}

void CNPUP_VisualManager::OnDrawRibbonMainPanelButtonBorder
(
	CDC * pDC,
	CMFCRibbonButton * pButton
)
{
	ColorTheme::ColorMainMenuRibbon clrMainMenuRibbon = m_currentTheme.m_ColorMainMenuRibbon;

	CRect rectButton = pButton->GetRect();

	COLORREF clrButtonBorder = RGB(255, 0, 0);

	BOOL buttonDisabled = pButton->IsDisabled();
	BOOL buttonHighlighted = pButton->IsHighlighted();
	BOOL buttonPressed = pButton->IsPressed();

	if (buttonDisabled)
	{
		clrButtonBorder = clrMainMenuRibbon.clr_RibbonMainMenuButtonBorder_Disabled;
	}
	else
	{
		if (buttonHighlighted)
		{
			clrButtonBorder = clrMainMenuRibbon.clr_RibbonMainMenuButtonBorder_Highlight;
		}
		else
		{
			if (buttonPressed)
			{
				clrButtonBorder = clrMainMenuRibbon.clr_RibbonMainMenuButtonBorder_Pressed;
			}
			else
			{
				clrButtonBorder = clrMainMenuRibbon.clr_RibbonMainMenuButtonBorder;
			}
		}
	}

	NVS_DrawRectBorder(pDC, rectButton, clrButtonBorder, 1, 2, 3, 4);

}
#pragma endregion

#pragma region OtherRBBN
COLORREF CNPUP_VisualManager::GetRibbonEditBackgroundColor
(
	CMFCRibbonRichEditCtrl * pEdit,
	BOOL bIsHighlighted,
	BOOL bIsPaneHighlighted,
	BOOL bIsDisabled
)
{
	ColorTheme::ColorControlElements clrCntrlElem = m_currentTheme.m_ColorContrlElem;

	COLORREF clrEdditBackground = clrCntrlElem.clr_RichEdit_Background;

	return clrEdditBackground;
}

void CNPUP_VisualManager::OnDrawRibbonMenuCheckFrame
(
	CDC * pDC,
	CMFCRibbonButton * pButton,
	CRect rect
)
{
	ColorTheme::ColorGlobalElement clrGlblElem = m_currentTheme.m_ColorGlobalElement;
	ColorTheme::ParametrsCMenuImage prmDraw = m_currentTheme.m_sParamCMenu;
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	CMenuImages currentMenuImages = m_CMIUserImage;

	COLORREF clrCheck = clrGlblElem.clr_CheckMenu;

	pDC->FillSolidRect(rect, clrCheck);

	currentMenuImages.Draw(pDC, CMenuImages::IdArrowDown, rect, clrImage);

}

void CNPUP_VisualManager::OnDrawRibbonSliderZoomButton
(
	CDC * pDC,
	CMFCRibbonSlider * pSlider,
	CRect rect,
	BOOL bIsZoomOut,
	BOOL bIsHighlighted,
	BOOL bIsPressed,
	BOOL bIsDisabled
)
{
	ColorTheme::ColorSliderZoom clrSlider = m_currentTheme.m_ColorSliderZoom;
	CMenuImages::IMAGE_STATE clrImage = m_currentTheme.m_sParamCMenu.img_ColorImage;
	CMenuImages currentMenuImages = m_CMIUserImage;

	COLORREF clrButton = RGB(255, 0, 0);

	CMFCRibbonCategory *pParentCategory = pSlider->GetParentCategory();
	AFX_RibbonCategoryColor colorContextTab = AFX_RibbonCategoryColor::AFX_CategoryColor_None;

	if (pParentCategory != NULL)
	{
		colorContextTab = pParentCategory->GetTabColor();
	}

	if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_None)
	{
		if (bIsDisabled)
		{
			if (bIsZoomOut)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Minus_Disabled;
			}
			else
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Plus_Disabled;
			}
		}
		else
		{
			if (bIsHighlighted)
			{
				if (bIsZoomOut)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Minus_Highlight;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Plus_Highlight;
				}
			}
			else
			{
				if (bIsPressed)
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Pressed;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Pressed;
					}
				}
				else
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus;
					}
				}
			}
		}
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Red)
	{
		if (bIsDisabled)
		{
			if (bIsZoomOut)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Minus_Disabled_Red;
			}
			else
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Plus_Disabled_Red;
			}
		}
		else
		{
			if (bIsHighlighted)
			{
				if (bIsZoomOut)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Minus_Highlight_Red;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Plus_Highlight_Red;
				}
			}
			else
			{
				if (bIsPressed)
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Pressed_Red;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Pressed_Red;
					}
				}
				else
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Red;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Red;
					}
				}
			}
		}
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Orange)
	{
		if (bIsDisabled)
		{
			if (bIsZoomOut)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Minus_Disabled_Orange;
			}
			else
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Plus_Disabled_Orange;
			}
		}
		else
		{
			if (bIsHighlighted)
			{
				if (bIsZoomOut)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Minus_Highlight_Orange;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Plus_Highlight_Orange;
				}
			}
			else
			{
				if (bIsPressed)
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Pressed_Orange;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Pressed_Orange;
					}
				}
				else
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Orange;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Orange;
					}
				}
			}
		}
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Yellow)
	{
		if (bIsDisabled)
		{
			if (bIsZoomOut)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Minus_Disabled_Yellow;
			}
			else
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Plus_Disabled_Yellow;
			}
		}
		else
		{
			if (bIsHighlighted)
			{
				if (bIsZoomOut)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Minus_Highlight_Yellow;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Plus_Highlight_Yellow;
				}
			}
			else
			{
				if (bIsPressed)
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Pressed_Yellow;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Pressed_Yellow;
					}
				}
				else
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Yellow;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Yellow;
					}
				}
			}
		}
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Green)
	{
		if (bIsDisabled)
		{
			if (bIsZoomOut)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Minus_Disabled_Green;
			}
			else
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Plus_Disabled_Green;
			}
		}
		else
		{
			if (bIsHighlighted)
			{
				if (bIsZoomOut)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Minus_Highlight_Green;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Plus_Highlight_Green;
				}
			}
			else
			{
				if (bIsPressed)
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Pressed_Green;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Pressed_Green;
					}
				}
				else
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Green;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Green;
					}
				}
			}
		}
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Blue)
	{
		if (bIsDisabled)
		{
			if (bIsZoomOut)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Minus_Disabled_Blue;
			}
			else
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Plus_Disabled_Blue;
			}
		}
		else
		{
			if (bIsHighlighted)
			{
				if (bIsZoomOut)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Minus_Highlight_Blue;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Plus_Highlight_Blue;
				}
			}
			else
			{
				if (bIsPressed)
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Pressed_Blue;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Pressed_Blue;
					}
				}
				else
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Blue;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Blue;
					}
				}
			}
		}
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Indigo)
	{
		if (bIsDisabled)
		{
			if (bIsZoomOut)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Minus_Disabled_Indigo;
			}
			else
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Plus_Disabled_Indigo;
			}
		}
		else
		{
			if (bIsHighlighted)
			{
				if (bIsZoomOut)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Minus_Highlight_Indigo;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Plus_Highlight_Indigo;
				}
			}
			else
			{
				if (bIsPressed)
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Pressed_Indigo;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Pressed_Indigo;
					}
				}
				else
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Indigo;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Indigo;
					}
				}
			}
		}
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Violet)
	{
		if (bIsDisabled)
		{
			if (bIsZoomOut)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Minus_Disabled_Violet;
			}
			else
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Plus_Disabled_Violet;
			}
		}
		else
		{
			if (bIsHighlighted)
			{
				if (bIsZoomOut)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Minus_Highlight_Violet;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Plus_Highlight_Violet;
				}
			}
			else
			{
				if (bIsPressed)
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Pressed_Violet;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Pressed_Violet;
					}
				}
				else
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Violet;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Violet;
					}
				}
			}
		}
	}
	else
	{
		if (bIsDisabled)
		{
			if (bIsZoomOut)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Minus_Disabled;
			}
			else
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Plus_Disabled;
			}
		}
		else
		{
			if (bIsHighlighted)
			{
				if (bIsZoomOut)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Minus_Highlight;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Plus_Highlight;
				}
			}
			else
			{
				if (bIsPressed)
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus_Pressed;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus_Pressed;
					}
				}
				else
				{
					if (bIsZoomOut)
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Minus;
					}
					else
					{
						clrButton = clrSlider.clr_BottonZoom_Background_Plus;
					}
				}
			}
		}
	}

	pDC->FillSolidRect(rect, clrButton);

	if (bIsZoomOut)
	{
		currentMenuImages.Draw(pDC, CMenuImages::IdArrowLeftLarge, rect, clrImage);
	}
	else
	{
		currentMenuImages.Draw(pDC, CMenuImages::IdArrowRightLarge, rect, clrImage);
	}

}

void CNPUP_VisualManager::OnDrawRibbonSliderChannel
(
	CDC * pDC,
	CMFCRibbonSlider * pSlider,
	CRect rect
)
{
	ColorTheme::ColorSliderZoom clrSlider = m_currentTheme.m_ColorSliderZoom;

	COLORREF clrButton = RGB(255, 0, 0);

	CMFCRibbonCategory *pParentCategory = pSlider->GetParentCategory();
	AFX_RibbonCategoryColor colorContextTab = AFX_RibbonCategoryColor::AFX_CategoryColor_None;

	if (pParentCategory != NULL)
	{
		colorContextTab = pParentCategory->GetTabColor();
	}

	if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_None)
	{
		clrButton = clrSlider.clr_SliderZoom_Background;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Red)
	{
		clrButton = clrSlider.clr_SliderZoom_Background_Red;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Orange)
	{
		clrButton = clrSlider.clr_SliderZoom_Background_Orange;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Yellow)
	{
		clrButton = clrSlider.clr_SliderZoom_Background_Yellow;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Green)
	{
		clrButton = clrSlider.clr_SliderZoom_Background_Green;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Blue)
	{
		clrButton = clrSlider.clr_SliderZoom_Background_Blue;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Indigo)
	{
		clrButton = clrSlider.clr_SliderZoom_Background_Indigo;
	}
	else
	{
		clrButton = clrSlider.clr_SliderZoom_Background;
	}

	pDC->FillSolidRect(rect, clrButton);
}

void CNPUP_VisualManager::OnDrawRibbonSliderThumb
(
	CDC * pDC,
	CMFCRibbonSlider * pSlider,
	CRect rect,
	BOOL bIsHighlighted,
	BOOL bIsPressed,
	BOOL bIsDisabled
)
{
	ColorTheme::ColorSliderZoom clrSlider = m_currentTheme.m_ColorSliderZoom;

	COLORREF clrButton = RGB(255, 0, 0);

	CMFCRibbonCategory *pParentCategory = pSlider->GetParentCategory();
	AFX_RibbonCategoryColor colorContextTab = AFX_RibbonCategoryColor::AFX_CategoryColor_None;

	if (pParentCategory != NULL)
	{
		colorContextTab = pParentCategory->GetTabColor();
	}

	if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_None)
	{
		if (bIsDisabled)
		{
			clrButton = clrSlider.clr_BottonZoom_Background_Disabled;
		}
		else
		{
			if (bIsPressed)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Pressed;
			}
			else
			{
				if (bIsHighlighted)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Highlight;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background;
				}
			}
		}
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Red)
	{
		if (bIsDisabled)
		{
			clrButton = clrSlider.clr_BottonZoom_Background_Disabled_Red;
		}
		else
		{
			if (bIsPressed)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Pressed_Red;
			}
			else
			{
				if (bIsHighlighted)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Highlight_Red;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Red;
				}
			}
		}
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Orange)
	{
		if (bIsDisabled)
		{
			clrButton = clrSlider.clr_BottonZoom_Background_Disabled_Orange;
		}
		else
		{
			if (bIsPressed)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Pressed_Orange;
			}
			else
			{
				if (bIsHighlighted)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Highlight_Orange;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Orange;
				}
			}
		}
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Yellow)
	{
		if (bIsDisabled)
		{
			clrButton = clrSlider.clr_BottonZoom_Background_Disabled_Yellow;
		}
		else
		{
			if (bIsPressed)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Pressed_Yellow;
			}
			else
			{
				if (bIsHighlighted)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Highlight_Yellow;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Yellow;
				}
			}
		}
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Green)
	{
		if (bIsDisabled)
		{
			clrButton = clrSlider.clr_BottonZoom_Background_Disabled_Green;
		}
		else
		{
			if (bIsPressed)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Pressed_Green;
			}
			else
			{
				if (bIsHighlighted)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Highlight_Green;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Green;
				}
			}
		}
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Blue)
	{
		if (bIsDisabled)
		{
			clrButton = clrSlider.clr_BottonZoom_Background_Disabled_Blue;
		}
		else
		{
			if (bIsPressed)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Pressed_Blue;
			}
			else
			{
				if (bIsHighlighted)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Highlight_Blue;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Blue;
				}
			}
		}
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Indigo)
	{
		if (bIsDisabled)
		{
			clrButton = clrSlider.clr_BottonZoom_Background_Disabled_Indigo;
		}
		else
		{
			if (bIsPressed)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Pressed_Indigo;
			}
			else
			{
				if (bIsHighlighted)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Highlight_Indigo;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Indigo;
				}
			}
		}
	}
	else
	{
		if (bIsDisabled)
		{
			clrButton = clrSlider.clr_BottonZoom_Background_Disabled;
		}
		else
		{
			if (bIsPressed)
			{
				clrButton = clrSlider.clr_BottonZoom_Background_Pressed;
			}
			else
			{
				if (bIsHighlighted)
				{
					clrButton = clrSlider.clr_BottonZoom_Background_Highlight;
				}
				else
				{
					clrButton = clrSlider.clr_BottonZoom_Background;
				}
			}
		}
	}

	pDC->FillSolidRect(rect, clrButton);

}

void CNPUP_VisualManager::OnDrawRibbonProgressBar
(
	CDC * pDC,
	CMFCRibbonProgressBar * pProgress,
	CRect rectProgress,
	CRect rectChunk,
	BOOL bInfiniteMode
)
{
	ColorTheme::ColorControlElements clrProgressBar = m_currentTheme.m_ColorContrlElem;

	COLORREF clrBar = RGB(255, 0, 0);
	COLORREF clrChunk = RGB(255, 0, 0);

	CMFCRibbonCategory *pParentCategory = pProgress->GetParentCategory();
	AFX_RibbonCategoryColor colorContextTab = AFX_RibbonCategoryColor::AFX_CategoryColor_None;

	if (pParentCategory != NULL)
	{
		colorContextTab = pParentCategory->GetTabColor();
	}

	if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_None)
	{
		clrBar = clrProgressBar.clr_ProgressBar_Background;
		clrChunk = clrProgressBar.clr_ProgressChunk_Background;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Red)
	{
		clrBar = clrProgressBar.clr_ProgressBar_Background_Red;
		clrChunk = clrProgressBar.clr_ProgressChunk_Background_Red;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Orange)
	{
		clrBar = clrProgressBar.clr_ProgressBar_Background_Orange;
		clrChunk = clrProgressBar.clr_ProgressChunk_Background_Orange;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Yellow)
	{
		clrBar = clrProgressBar.clr_ProgressBar_Background_Yellow;
		clrChunk = clrProgressBar.clr_ProgressChunk_Background_Yellow;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Green)
	{
		clrBar = clrProgressBar.clr_ProgressBar_Background_Green;
		clrChunk = clrProgressBar.clr_ProgressChunk_Background_Green;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Blue)
	{
		clrBar = clrProgressBar.clr_ProgressBar_Background_Blue;
		clrChunk = clrProgressBar.clr_ProgressChunk_Background_Blue;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Indigo)
	{
		clrBar = clrProgressBar.clr_ProgressBar_Background_Indigo;
		clrChunk = clrProgressBar.clr_ProgressChunk_Background_Indigo;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Violet)
	{
		clrBar = clrProgressBar.clr_ProgressBar_Background_Violet;
		clrChunk = clrProgressBar.clr_ProgressChunk_Background_Violet;
	}
	else
	{
		clrBar = clrProgressBar.clr_ProgressBar_Background;
		clrChunk = clrProgressBar.clr_ProgressChunk_Background;
	}

	pDC->FillSolidRect(rectProgress, clrBar); //фон
	pDC->FillSolidRect(rectChunk, clrChunk); //сам бар
}

COLORREF CNPUP_VisualManager::GetRibbonHyperlinkTextColor
(
	CMFCRibbonLinkCtrl * pHyperLink
)
{
	ColorTheme::ColorControlElements clrCntrlElem = m_currentTheme.m_ColorContrlElem;

	COLORREF clrText = RGB(255, 0, 0);

	CMFCRibbonCategory *pParentCategory = pHyperLink->GetParentCategory();
	AFX_RibbonCategoryColor colorContextTab = AFX_RibbonCategoryColor::AFX_CategoryColor_None;

	if (pParentCategory != NULL)
	{
		colorContextTab = pParentCategory->GetTabColor();
	}

	if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_None)
	{
		clrText = clrCntrlElem.clr_TextHiperLink;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Red)
	{
		clrText = clrCntrlElem.clr_TextHiperLink_Red;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Orange)
	{
		clrText = clrCntrlElem.clr_TextHiperLink_Orange;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Yellow)
	{
		clrText = clrCntrlElem.clr_TextHiperLink_Yellow;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Green)
	{
		clrText = clrCntrlElem.clr_TextHiperLink_Green;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Blue)
	{
		clrText = clrCntrlElem.clr_TextHiperLink_Blue;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Indigo)
	{
		clrText = clrCntrlElem.clr_TextHiperLink_Indigo;
	}
	else if (colorContextTab == AFX_RibbonCategoryColor::AFX_CategoryColor_Violet)
	{
		clrText = clrCntrlElem.clr_TextHiperLink_Violet;
	}
	else
	{
		clrText = clrCntrlElem.clr_TextHiperLink;
	}

	return clrText;
}

void CNPUP_VisualManager::OnDrawRibbonLabel
(
	CDC * pDC,
	CMFCRibbonLabel * pLabel,
	CRect rect
)
{
	ColorTheme::ColorRibbonPanel clrRibbonPanel = m_currentTheme.m_ColorRibbonPanel;
	ColorTheme::ParametrsDrawRibbon prmDraw = m_currentTheme.m_sParamRibbon;

	COLORREF colorBackgroundCaptionPanel = RGB(0, 0, 0);
	COLORREF colorBorderCaptionPanel = RGB(0, 0, 0);

	CMFCRibbonCategory *categoryParrent = NULL;
	AFX_RibbonCategoryColor colorCategoryTab = AFX_CategoryColor_None;

	BOOL bIsHighlighted = pLabel->IsHighlighted(); // потом

	categoryParrent = pLabel->GetParentCategory();
	if (categoryParrent == NULL)
	{
		return;
	}

	colorCategoryTab = categoryParrent->GetTabColor();

		if (colorCategoryTab == AFX_CategoryColor_None)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Red)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Red;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Orange)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Orange;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Yellow)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Yellow;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Green)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Green;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Blue)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Blue;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Indigo)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Indigo;
		}
		else if (colorCategoryTab == AFX_CategoryColor_Violet)
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground_Violet;
		}
		else
		{
			colorBackgroundCaptionPanel = clrRibbonPanel.clr_RibbonPanelTabLabelBackground;
		}

	pDC->FillSolidRect(rect, colorBackgroundCaptionPanel);


}

#pragma endregion

#pragma endregion


