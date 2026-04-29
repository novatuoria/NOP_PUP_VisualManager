// ====================== NOVATUORIA   ORIGINAL   PROJECT =====================
//
//
//                //|| //  //==//  //==//    //==//  //  //  //==//
//               // ||//  //  //  //==//    //==//  //  //  //==//
//              //  |//  //==//  //        //      //==//  //
//
//
//            NOP PUP Visual Manager for MFC (ver. 0.1 "helloWorld!")           
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
// Привет! :)
// Это альфа версия NOP PUP Visual Manager. Она содержит функции которые были 
// проверены и чей функционал известен и документирован более чем наполовину.
// Здесь нет экспериментальных функций, но есть функции чья взаимосвязь и вся
// техническая функциональность не изучена до конца по причине невозможности
// провести полное исследование в режиме одиночной разработки без качественного 
// тестирования в условиях реальной эксплуатации.
//
// ----------------------------------------------------------------------------
//
// Hi! :)
// This is a alpha version of the Visual Manager. It contains functions that 
// have been verified, but their documentation and full functional scope 
// are more than 50% complete.
// While it excludes purely experimental features, certain internal 
// dependencies and technical behaviors are not yet fully explored. 
// This is due to the limitations of solo development and the lack of 
// large-scale testing in real-world production environments.
//
// ============================================================================
//
// [ Технический обзор ]
// * Полная визуальная переработка интерфейсов на базе MFC (Риббон, панели, 
// меню, диалоги).
// * Централизованное управление цветами через структуру конфигурации тем.
// * Кастомная отрисовка с использованием GDI.
//
// ----------------------------------------------------------------------------
//
// [ Тестирование ]
// В процессе разработки и отладки использовались официальные примеры 
// Microsoft (VCSamples: MSOffice2007Demo, RibbonGadgets и др.). 
// Это позволило обеспечить полную нативную совместимость с библиотекой MFC.
//
// ----------------------------------------------------------------------------
//
// [ Как подключить ]
// 1. Подключите заголовочный файл в месте инициализации Visual Manager:
//    #include "NOP_PUP_VisualManager.h"
//
// 2. Установите желаемый стиль одной строкой (например, в InitInstance):
//    CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_Blue);
// 
// 3. Вызовите метод: 
// CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));
//
// Поддержка Runtime: Смена тем происходит мгновенно без перезагрузки 
// интерфейса. Достаточно повторного вызова SetStyle() и SetDefaultManager();
//
// ----------------------------------------------------------------------------
//
// [ Структура файла / File Structure ]
// Данный файл состоит из следующих регионов, которые группируют методы 
// визуального менеджера (this file is organized into the following regions, 
// grouping the visual manager's methods) :
//
//		Region StandardColor
//
//		Region RibbonColor
//
//		Region Standard
//
//		Region Ribbon
//



#pragma once
#include <afxautohidedocksite.h>
#include <afxvisualmanageroffice2007.h>


class CNPUP_VisualManager;


class CNPUP_VisualManager : public CMFCVisualManagerOfficeXP
{

	DECLARE_DYNCREATE(CNPUP_VisualManager)

public:


	struct DefaultColorTheme
	{
		// ----------------------------------------------------------------------------
		// ГЛОБАЛЬНЫЕ ПАЛИТРЫ ТЕМ / GLOBAL THEME PALETS
		// ----------------------------------------------------------------------------

		COLORREF clr_ClobalBackground;
		COLORREF clr_ClientFrameBackground;
		COLORREF clr_ClientBorder;

		COLORREF clr_Caption;
		COLORREF clr_ActiveCaption;
		COLORREF clr_DisableCaption;

		COLORREF clr_ButtonClient;
		COLORREF clr_ActiveButtonClient;
		COLORREF clr_DisableButtonClient;
		COLORREF clr_PressedButtonClient;

		COLORREF clr_SysButton;
		COLORREF clr_ActiveSysButton;
		COLORREF clr_DisableSysButton;
		COLORREF clr_PressedSysButton;
		COLORREF clr_HighlightSysButton;

		COLORREF clr_BorderButtonClient;
		COLORREF clr_BorderActiveButtonClient;
		COLORREF clr_BorderDisableButtonClient;
		COLORREF clr_BorderPressedButtonClient;

		COLORREF clr_BorderSysButton;
		COLORREF clr_BorderActiveSysButton;
		COLORREF clr_BorderDisableSysButton;
		COLORREF clr_BorderPressedSysButton;
		COLORREF clr_BorderHighlightSysButton;

		COLORREF clr_RibbonFrame;
		COLORREF clr_RibbonFrame_Red;
		COLORREF clr_RibbonFrame_Orange;
		COLORREF clr_RibbonFrame_Yellow;
		COLORREF clr_RibbonFrame_Green;
		COLORREF clr_RibbonFrame_Blue;
		COLORREF clr_RibbonFrame_Indigo;
		COLORREF clr_RibbonFrame_Violet;

		COLORREF clr_RibbonFrame_Border;
		COLORREF clr_RibbonFrame_Border_Red;
		COLORREF clr_RibbonFrame_Border_Orange;
		COLORREF clr_RibbonFrame_Border_Yellow;
		COLORREF clr_RibbonFrame_Border_Green;
		COLORREF clr_RibbonFrame_Border_Blue;
		COLORREF clr_RibbonFrame_Border_Indigo;
		COLORREF clr_RibbonFrame_Border_Violet;

		COLORREF clr_ImageIcon;

		COLORREF clr_Text;
		COLORREF clr_TextActive;
		COLORREF clr_TextDisable;
		COLORREF clr_TextPressed;
		COLORREF clr_TextHyperlink;

	};

	struct ColorTheme
	{

		// ----------------------------------------------------------------------------
		// КЛАССИЧЕСКИЕ ЭЛЕМЕНТЫ ИНТЕРФЕЙСА / STANDARD UI FUNCTIONS
		// ----------------------------------------------------------------------------
		struct ColorGlobalElement
		{
			COLORREF clr_BackgroundFrame;							
			COLORREF clr_BackgroundClientFrame;						
			COLORREF clr_DividerFrame;								
			COLORREF clr_BorderFrame;								

			COLORREF clr_BackgroundMenu;								
			COLORREF clr_BackgroundLabelMenu;							
			COLORREF clr_BorderMenu;									
			COLORREF clr_Separator;									
			COLORREF clr_SelectedMenu;								
			COLORREF clr_ImportantMenu;								

			COLORREF clr_CheckMenu;									
			COLORREF clr_ResizeBar;									
			COLORREF clr_ShowAllMenuBackground;						
			COLORREF clr_ShowAllMenuBackground_Pressed;				
			COLORREF clr_ShowAllMenuBackground_Highlight;			

			COLORREF clr_CheckBoxBackground;					
			COLORREF clr_CheckBoxBackground_Enable;				
			COLORREF clr_CheckBoxBackground_Disable;			
			COLORREF clr_CheckBoxBackground_Highlight;			
			COLORREF clr_CheckBoxBackground_Pressed;			

			COLORREF clr_CheckBoxBorder;					
			COLORREF clr_CheckBoxBorder_Enable;					
			COLORREF clr_CheckBoxBorder_Disable;			
			COLORREF clr_CheckBoxBorder_Highlight;			
			COLORREF clr_CheckBoxBorder_Pressed;		

			COLORREF clr_SpinButtonsBackground_Up;				
			COLORREF clr_SpinButtonsBackground_Down;			
			COLORREF clr_SpinButtonsBackgroundPressed_Up;		
			COLORREF clr_SpinButtonsBackgroundPressed_Down;		
			COLORREF clr_SpinButtonsBackgroundHighlight_Up;		
			COLORREF clr_SpinButtonsBackgroundHighlight_Down;	
			COLORREF clr_SpinButtonsBackgroundDisable_Up;			
			COLORREF clr_SpinButtonsBackgroundDisable_Down;		

			COLORREF clr_SpinButtonsBorder_Up;					
			COLORREF clr_SpinButtonsBorder_Down;			
			COLORREF clr_SpinButtonsBorderPressed_Up;		
			COLORREF clr_SpinButtonsBorderPressed_Down;		
			COLORREF clr_SpinButtonsBorderHighlight_Up;		
			COLORREF clr_SpinButtonsBorderHighlight_Down;	
			COLORREF clr_SpinButtonsBorderDisable_Up;		
			COLORREF clr_SpinButtonsBorderDisable_Down;			

			COLORREF clr_ToolTipInfoBackground;	
			COLORREF clr_ToolTipInfoGradient;			
			COLORREF clr_ToolTipInfoBorder;				

			COLORREF clr_CommandListBackground;			
			COLORREF clr_CommandListBackground_Selected;	
			COLORREF clr_CommandListBackground_Highlight;	

			COLORREF clr_TextFrame;
			COLORREF clr_TextLabelMenu;					
			COLORREF clr_TextMenu;								
			COLORREF clr_TextMenu_Selected;					
			COLORREF clr_TextMenu_Pressed;						
			COLORREF clr_TextMenu_Disable;				
			COLORREF clr_TextMenu_Highlight;			
			COLORREF clr_ToolTipInfoText;				
			COLORREF clr_TextCommandList;				
			COLORREF clr_TextCommandList_Selected;		
			COLORREF clr_TextCommandList_Highlight;			
		};

		struct ColorMiniFrame
		{
			COLORREF clr_BackgroundMiniFrame;					
			COLORREF clr_BackgroundMiniFrame_CaptionActive;	
			COLORREF clr_BackgroundMiniFrame_CaptionDisable;	
			COLORREF clr_BackgroundMiniFrame_BorderActive;	
			COLORREF clr_BackgroundMiniFrame_BorderDisable;	
			COLORREF clr_BackgroundMiniFrame_Client;			

			COLORREF clr_BorderMiniFrame_Client;		

			COLORREF clr_ButtonMiniFrame;			
			COLORREF clr_ButtonMiniFrame_Active;				
			COLORREF clr_ButtonMiniFrame_Disable;			
			COLORREF clr_ButtonMiniFrame_Highlight;				

			COLORREF clr_ButtonBorderMiniFrame_Active;	
			COLORREF clr_ButtonBorderMiniFrame_Disable;		
			COLORREF clr_ButtonBorderMiniFrame_Highlight;		

			COLORREF clr_TextCaptionMiniFrame;				
			COLORREF clr_TextCaptionMiniFrame_Active;		
			COLORREF clr_TextCaptionMiniFrame_Disable;	

			COLORREF clr_ButtonsElement;				
			COLORREF clr_ButtonsElement_Active;			
			COLORREF clr_ButtonsElement_Disable;			
			COLORREF clr_ButtonsElement_Dropped;			
			COLORREF clr_ButtonsElement_NoDropped;		
			COLORREF clr_ButtonsElement_Highlight;	

			COLORREF clr_ButtonsBorderElement;			
			COLORREF clr_ButtonsBorderElement_Active;	
			COLORREF clr_ButtonsBorderElement_Disable;		
			COLORREF clr_ButtonsBorderElement_Dropped;		
			COLORREF clr_ButtonsBorderElement_NoDropped;	
			COLORREF clr_ButtonsBorderElement_Highlight;	
		};

		struct ColorTabs
		{
			COLORREF clr_BackgroundFrameDockMDI;				
			COLORREF clr_BackgroundFrameDockHide;			

			COLORREF clr_BackgroundTabsHide;			
			COLORREF clr_BackgroundTabsHide_Active;				
			COLORREF clr_BackgroundTabsHide_Disable;		
			COLORREF clr_BackgroundTabsHide_Highlight;		

			COLORREF clr_BackgroundTabsMDI;				
			COLORREF clr_BackgroundTabsMDI_Active;			
			COLORREF clr_BackgroundTabsMDI_Disable;		
			COLORREF clr_BackgroundTabsMDI_Highlight;			

			COLORREF clr_BorderTabsMDI;							
			COLORREF clr_BorderTabsMDI_Active;					
			COLORREF clr_BorderTabsMDI_Disable;					
			COLORREF clr_BorderTabsMDI_Highlight;				

			COLORREF clr_BorderTabsHide;							
			COLORREF clr_BorderTabsHide_Active;						
			COLORREF clr_BorderTabsHide_Disable;						
			COLORREF clr_BorderTabsHide_Highlight;			

			COLORREF clr_ButtonTabsMDI;							
			COLORREF clr_ButtonTabsMDI_Active;						
			COLORREF clr_ButtonTabsMDI_Disable;					
			COLORREF clr_ButtonTabsMDI_Highlight;				

			COLORREF clr_BorderButtonTabsMDI;							
			COLORREF clr_BorderButtonTabsMDI_Active;				
			COLORREF clr_BorderButtonTabsMDI_Disable;					
			COLORREF clr_BorderButtonTabsMDI_Highlight;					

			COLORREF clr_TextTabsMDI;									
			COLORREF clr_TextTabsMDI_Active;							
			COLORREF clr_TextTabsMDI_Disable;							
			COLORREF clr_TextTabsMDI_Highlight;						
			
			COLORREF clr_TextTabsHide;										
			COLORREF clr_TextTabsHide_Active;								
			COLORREF clr_TextTabsHide_Disable;								
			COLORREF clr_TextTabsHide_Highlight;							
		};

		struct ColorPropertyGrid
		{
			COLORREF clr_PropGrid_Background;						
			COLORREF clr_PropGrid_GroupBackground;					
			COLORREF clr_PropGrid_DescriptionBackground;			
			COLORREF clr_PropGrid_Line;								

			COLORREF clr_PropGrid_Text;									
			COLORREF clr_PropGrid_GroupText;							
			COLORREF clr_PropGrid_DescriptionText;						
		};

		struct ColorThreeView
		{
			COLORREF clr_ThreeViewBackground;
			COLORREF clr_ThreeViewLine;

			COLORREF clr_ThreeView_Text;									

		};

		struct ColorToolbars
		{
			COLORREF clr_BackgroundToolbar;
			COLORREF clr_BorderToolbar;

			COLORREF clr_BackgroundIconToolbar;
			COLORREF clr_BackgroundIconToolbar_Pressed;
			COLORREF clr_BackgroundIconToolbar_Highlight;

			COLORREF clr_BorderIconToolbar;
			COLORREF clr_BorderIconToolbar_Pressed;
			COLORREF clr_BorderIconToolbar_Highlight;

			COLORREF clr_ToolbarGripper;

			COLORREF clr_ToolbarText;
			COLORREF clr_ToolbarText_Disable;
			COLORREF clr_ToolbarText_Highlight;
			COLORREF clr_ToolbarText_Pressed;
		};

		struct ColorPaneCaptionsAndButton
		{
			COLORREF clr_PaneCaption_BackgroundFrame;
			COLORREF clr_PaneCaption_BackgroundCaption;
			COLORREF clr_PaneCaption_Border;

			COLORREF clr_PaneCaption_ButtonsBackground;
			COLORREF clr_PaneCaption_ButtonsBackground_Disable;
			COLORREF clr_PaneCaption_ButtonsBackground_Pressed;
			COLORREF clr_PaneCaption_ButtonsBackground_Highlight;

			COLORREF clr_PaneCaption_ButtonsBorder;
			COLORREF clr_PaneCaption_ButtonsBorder_Disable;
			COLORREF clr_PaneCaption_ButtonsBorder_Pressed;
			COLORREF clr_PaneCaption_ButtonsBorder_Highlight;

			COLORREF clr_PaneCaption_Text;
			COLORREF clr_PaneCaption_ButtonsText;
			COLORREF clr_PaneCaption_ButtonsText_Disable;
			COLORREF clr_PaneCaption_ButtonsText_Pressed;
			COLORREF clr_PaneCaption_ButtonsText_Highlight;
		};

		struct ColorOutlookBar
		{
			COLORREF clr_OutlookBarCaptionBackground;
			COLORREF clr_OutlookBarSeparator;

			COLORREF clr_OutlookBarButtonBackground;
			COLORREF clr_OutlookBarButtonBackground_Pressed;
			COLORREF clr_OutlookBarButtonBackground_Highlight;

			COLORREF clr_OutlookBarButtonBorder;
			COLORREF clr_OutlookBarButtonBorder_Pressed;
			COLORREF clr_OutlookBarButtonBorder_Highlight;

			COLORREF clr_OutlookBarCaptionText;
			COLORREF clr_OutlookBarButtonText;
			COLORREF clr_OutlookBarButtonText_Pressed;
			COLORREF clr_OutlookBarButtonText_Highlight;
		};

		// ----------------------------------------------------------------------------
		// ЭЛЕМЕНТЫ ЛЕНТОЧНОГО ИНТЕРФЕЙСА / RIBBON UI FUNCTIONS
		// ----------------------------------------------------------------------------
		struct ColorRibbonBackground
		{
			COLORREF clr_RibbonBackground;

			COLORREF clr_RibbonCategoryCaption;
			COLORREF clr_RibbonCategoryCaption_Red;
			COLORREF clr_RibbonCategoryCaption_Orange;
			COLORREF clr_RibbonCategoryCaption_Yellow;
			COLORREF clr_RibbonCategoryCaption_Green;
			COLORREF clr_RibbonCategoryCaption_Blue;
			COLORREF clr_RibbonCategoryCaption_Indigo;
			COLORREF clr_RibbonCategoryCaption_Violet;

			COLORREF clr_RibbonCategoryCaption_Highlight;
			COLORREF clr_RibbonCategoryCaption_Highlight_Red;
			COLORREF clr_RibbonCategoryCaption_Highlight_Orange;
			COLORREF clr_RibbonCategoryCaption_Highlight_Yellow;
			COLORREF clr_RibbonCategoryCaption_Highlight_Green;
			COLORREF clr_RibbonCategoryCaption_Highlight_Blue;
			COLORREF clr_RibbonCategoryCaption_Highlight_Indigo;
			COLORREF clr_RibbonCategoryCaption_Highlight_Violet;

			COLORREF clr_RibbonCategoryCaption_Active;
			COLORREF clr_RibbonCategoryCaption_Active_Red;
			COLORREF clr_RibbonCategoryCaption_Active_Orange;
			COLORREF clr_RibbonCategoryCaption_Active_Yellow;
			COLORREF clr_RibbonCategoryCaption_Active_Green;
			COLORREF clr_RibbonCategoryCaption_Active_Blue;
			COLORREF clr_RibbonCategoryCaption_Active_Indigo;
			COLORREF clr_RibbonCategoryCaption_Active_Violet;

			COLORREF clr_RibbonCategoryCaptionBorder;
			COLORREF clr_RibbonCategoryCaptionBorder_Red;
			COLORREF clr_RibbonCategoryCaptionBorder_Orange;
			COLORREF clr_RibbonCategoryCaptionBorder_Yellow;
			COLORREF clr_RibbonCategoryCaptionBorder_Green;
			COLORREF clr_RibbonCategoryCaptionBorder_Blue;
			COLORREF clr_RibbonCategoryCaptionBorder_Indigo;
			COLORREF clr_RibbonCategoryCaptionBorder_Violet;

			COLORREF clr_RibbonCategoryCaptionBorder_Highlight;
			COLORREF clr_RibbonCategoryCaptionBorder_Highlight_Red;
			COLORREF clr_RibbonCategoryCaptionBorder_Highlight_Orange;
			COLORREF clr_RibbonCategoryCaptionBorder_Highlight_Yellow;
			COLORREF clr_RibbonCategoryCaptionBorder_Highlight_Green;
			COLORREF clr_RibbonCategoryCaptionBorder_Highlight_Blue;
			COLORREF clr_RibbonCategoryCaptionBorder_Highlight_Indigo;
			COLORREF clr_RibbonCategoryCaptionBorder_Highlight_Violet;
		};

		struct ColorMainMenuRibbon
		{
			COLORREF clr_RibbonMainMenuApplicationButtonBackground;
			COLORREF clr_RibbonMainMenuApplicationButtonBackground_Disabled;
			COLORREF clr_RibbonMainMenuApplicationButtonBackground_Highlight;
			COLORREF clr_RibbonMainMenuApplicationButtonBackground_Pressed;

			COLORREF clr_RibbonMainMenuMainPanelBackground;

			COLORREF clr_RibbonMainMenuButtonBackground;
			COLORREF clr_RibbonMainMenuButtonBackground_Disabled;
			COLORREF clr_RibbonMainMenuButtonBackground_Highlight;
			COLORREF clr_RibbonMainMenuButtonBackground_Pressed;

			COLORREF clr_RibbonMainMenuButtonBorder;
			COLORREF clr_RibbonMainMenuButtonBorder_Disabled;
			COLORREF clr_RibbonMainMenuButtonBorder_Highlight;
			COLORREF clr_RibbonMainMenuButtonBorder_Pressed;

			COLORREF clr_RibbonMainMenuButtonText;
			COLORREF clr_RibbonMainMenuButtonText_Disabled;
			COLORREF clr_RibbonMainMenuButtonText_Highlight;
			COLORREF clr_RibbonMainMenuButtonText_Pressed;
		};

		struct ColorRibbonTabs
		{
			COLORREF clr_RibbonTabs;
			COLORREF clr_RibbonTabs_Red;
			COLORREF clr_RibbonTabs_Orange;
			COLORREF clr_RibbonTabs_Yellow;
			COLORREF clr_RibbonTabs_Green;
			COLORREF clr_RibbonTabs_Blue;
			COLORREF clr_RibbonTabs_Indigo;
			COLORREF clr_RibbonTabs_Violet;

			COLORREF clr_RibbonTabs_Active;
			COLORREF clr_RibbonTabs_Active_Red;
			COLORREF clr_RibbonTabs_Active_Orange;
			COLORREF clr_RibbonTabs_Active_Yellow;
			COLORREF clr_RibbonTabs_Active_Green;
			COLORREF clr_RibbonTabs_Active_Blue;
			COLORREF clr_RibbonTabs_Active_Indigo;
			COLORREF clr_RibbonTabs_Active_Violet;

			COLORREF clr_RibbonTabs_Highlight;
			COLORREF clr_RibbonTabs_Highlight_Red;
			COLORREF clr_RibbonTabs_Highlight_Orange;
			COLORREF clr_RibbonTabs_Highlight_Yellow;
			COLORREF clr_RibbonTabs_Highlight_Green;
			COLORREF clr_RibbonTabs_Highlight_Blue;
			COLORREF clr_RibbonTabs_Highlight_Indigo;
			COLORREF clr_RibbonTabs_Highlight_Violet;

			COLORREF clr_RibbonTabsBorder;
			COLORREF clr_RibbonTabsBorder_Red;
			COLORREF clr_RibbonTabsBorder_Orange;
			COLORREF clr_RibbonTabsBorder_Yellow;
			COLORREF clr_RibbonTabsBorder_Green;
			COLORREF clr_RibbonTabsBorder_Blue;
			COLORREF clr_RibbonTabsBorder_Indigo;
			COLORREF clr_RibbonTabsBorder_Violet;

			COLORREF clr_RibbonTabsBorder_Active;
			COLORREF clr_RibbonTabsBorder_Active_Red;
			COLORREF clr_RibbonTabsBorder_Active_Orange;
			COLORREF clr_RibbonTabsBorder_Active_Yellow;
			COLORREF clr_RibbonTabsBorder_Active_Green;
			COLORREF clr_RibbonTabsBorder_Active_Blue;
			COLORREF clr_RibbonTabsBorder_Active_Indigo;
			COLORREF clr_RibbonTabsBorder_Active_Violet;

			COLORREF clr_RibbonTabsBorder_Highlight;
			COLORREF clr_RibbonTabsBorder_Highlight_Red;
			COLORREF clr_RibbonTabsBorder_Highlight_Orange;
			COLORREF clr_RibbonTabsBorder_Highlight_Yellow;
			COLORREF clr_RibbonTabsBorder_Highlight_Green;
			COLORREF clr_RibbonTabsBorder_Highlight_Blue;
			COLORREF clr_RibbonTabsBorder_Highlight_Indigo;
			COLORREF clr_RibbonTabsBorder_Highlight_Violet;

			COLORREF clr_RibbonTabs_Text;
			COLORREF clr_RibbonTabs_Text_Red;
			COLORREF clr_RibbonTabs_Text_Orange;
			COLORREF clr_RibbonTabs_Text_Yellow;
			COLORREF clr_RibbonTabs_Text_Green;
			COLORREF clr_RibbonTabs_Text_Blue;
			COLORREF clr_RibbonTabs_Text_Indigo;
			COLORREF clr_RibbonTabs_Text_Violet;

			COLORREF clr_RibbonTabs_Text_Active;
			COLORREF clr_RibbonTabs_Text_Active_Red;
			COLORREF clr_RibbonTabs_Text_Active_Orange;
			COLORREF clr_RibbonTabs_Text_Active_Yellow;
			COLORREF clr_RibbonTabs_Text_Active_Green;
			COLORREF clr_RibbonTabs_Text_Active_Blue;
			COLORREF clr_RibbonTabs_Text_Active_Indigo;
			COLORREF clr_RibbonTabs_Text_Active_Violet;

			COLORREF clr_RibbonTabs_Text_Highlight;
			COLORREF clr_RibbonTabs_Text_Highlight_Red;
			COLORREF clr_RibbonTabs_Text_Highlight_Orange;
			COLORREF clr_RibbonTabs_Text_Highlight_Yellow;
			COLORREF clr_RibbonTabs_Text_Highlight_Green;
			COLORREF clr_RibbonTabs_Text_Highlight_Blue;
			COLORREF clr_RibbonTabs_Text_Highlight_Indigo;
			COLORREF clr_RibbonTabs_Text_Highlight_Violet;
		};

		struct ColorRibbonCategory
		{
			COLORREF clr_RibbonCategoryBackground;
			COLORREF clr_RibbonCategoryBackground_Red;
			COLORREF clr_RibbonCategoryBackground_Orange;
			COLORREF clr_RibbonCategoryBackground_Yellow;
			COLORREF clr_RibbonCategoryBackground_Green;
			COLORREF clr_RibbonCategoryBackground_Blue;
			COLORREF clr_RibbonCategoryBackground_Indigo;
			COLORREF clr_RibbonCategoryBackground_Violet;

			COLORREF clr_RibbonCategoryBorder;
			COLORREF clr_RibbonCategoryBorder_Red;
			COLORREF clr_RibbonCategoryBorder_Orange;
			COLORREF clr_RibbonCategoryBorder_Yellow;
			COLORREF clr_RibbonCategoryBorder_Green;
			COLORREF clr_RibbonCategoryBorder_Blue;
			COLORREF clr_RibbonCategoryBorder_Indigo;
			COLORREF clr_RibbonCategoryBorder_Violet;
		};

		struct ColorRibbonPanel
		{
			COLORREF clr_RibbonPanelTabBackground;
			COLORREF clr_RibbonPanelTabBackground_Red;
			COLORREF clr_RibbonPanelTabBackground_Orange;
			COLORREF clr_RibbonPanelTabBackground_Yellow;
			COLORREF clr_RibbonPanelTabBackground_Green;
			COLORREF clr_RibbonPanelTabBackground_Blue;
			COLORREF clr_RibbonPanelTabBackground_Indigo;
			COLORREF clr_RibbonPanelTabBackground_Violet;

			COLORREF clr_RibbonPanelTabBackground_Highlight;
			COLORREF clr_RibbonPanelTabBackground_Highlight_Red;
			COLORREF clr_RibbonPanelTabBackground_Highlight_Orange;
			COLORREF clr_RibbonPanelTabBackground_Highlight_Yellow;
			COLORREF clr_RibbonPanelTabBackground_Highlight_Green;
			COLORREF clr_RibbonPanelTabBackground_Highlight_Blue;
			COLORREF clr_RibbonPanelTabBackground_Highlight_Indigo;
			COLORREF clr_RibbonPanelTabBackground_Highlight_Violet;

			COLORREF clr_RibbonPanelTabLabelBackground;
			COLORREF clr_RibbonPanelTabLabelBackground_Red;
			COLORREF clr_RibbonPanelTabLabelBackground_Orange;
			COLORREF clr_RibbonPanelTabLabelBackground_Yellow;
			COLORREF clr_RibbonPanelTabLabelBackground_Green;
			COLORREF clr_RibbonPanelTabLabelBackground_Blue;
			COLORREF clr_RibbonPanelTabLabelBackground_Indigo;
			COLORREF clr_RibbonPanelTabLabelBackground_Violet;

			COLORREF clr_RibbonPanelTabLabelBorder;
			COLORREF clr_RibbonPanelTabLabelBorder_Red;
			COLORREF clr_RibbonPanelTabLabelBorder_Orange;
			COLORREF clr_RibbonPanelTabLabelBorder_Yellow;
			COLORREF clr_RibbonPanelTabLabelBorder_Green;
			COLORREF clr_RibbonPanelTabLabelBorder_Blue;
			COLORREF clr_RibbonPanelTabLabelBorder_Indigo;
			COLORREF clr_RibbonPanelTabLabelBorder_Violet;

			COLORREF clr_RibbonPanelTabBorder;
			COLORREF clr_RibbonPanelTabBorder_Red;
			COLORREF clr_RibbonPanelTabBorder_Orange;
			COLORREF clr_RibbonPanelTabBorder_Yellow;
			COLORREF clr_RibbonPanelTabBorder_Green;
			COLORREF clr_RibbonPanelTabBorder_Blue;
			COLORREF clr_RibbonPanelTabBorder_Indigo;
			COLORREF clr_RibbonPanelTabBorder_Violet;

			COLORREF clr_RibbonPanelTabBorder_Highlight;
			COLORREF clr_RibbonPanelTabBorder_Highlight_Red;
			COLORREF clr_RibbonPanelTabBorder_Highlight_Orange;
			COLORREF clr_RibbonPanelTabBorder_Highlight_Yellow;
			COLORREF clr_RibbonPanelTabBorder_Highlight_Green;
			COLORREF clr_RibbonPanelTabBorder_Highlight_Blue;
			COLORREF clr_RibbonPanelTabBorder_Highlight_Indigo;
			COLORREF clr_RibbonPanelTabBorder_Highlight_Violet;

			COLORREF clr_RibbonPanelTabLabelText;
			COLORREF clr_RibbonPanelTabLabelText_Red;
			COLORREF clr_RibbonPanelTabLabelText_Orange;
			COLORREF clr_RibbonPanelTabLabelText_Yellow;
			COLORREF clr_RibbonPanelTabLabelText_Green;
			COLORREF clr_RibbonPanelTabLabelText_Blue;
			COLORREF clr_RibbonPanelTabLabelText_Indigo;
			COLORREF clr_RibbonPanelTabLabelText_Violet;

			COLORREF clr_RibbonPanelTabLabelText_Highlight;
			COLORREF clr_RibbonPanelTabLabelText_Highlight_Red;
			COLORREF clr_RibbonPanelTabLabelText_Highlight_Orange;
			COLORREF clr_RibbonPanelTabLabelText_Highlight_Yellow;
			COLORREF clr_RibbonPanelTabLabelText_Highlight_Green;
			COLORREF clr_RibbonPanelTabLabelText_Highlight_Blue;
			COLORREF clr_RibbonPanelTabLabelText_Highlight_Indigo;
			COLORREF clr_RibbonPanelTabLabelText_Highlight_Violet;

		};

		struct ColorRibbonButton
		{
			COLORREF clr_RibbonButtonBackground;
			COLORREF clr_RibbonButtonBackground_Red;
			COLORREF clr_RibbonButtonBackground_Orange;
			COLORREF clr_RibbonButtonBackground_Yellow;
			COLORREF clr_RibbonButtonBackground_Green;
			COLORREF clr_RibbonButtonBackground_Blue;
			COLORREF clr_RibbonButtonBackground_Indigo;
			COLORREF clr_RibbonButtonBackground_Violet;

			COLORREF clr_RibbonButtonBackground_Disable;
			COLORREF clr_RibbonButtonBackground_Disable_Red;
			COLORREF clr_RibbonButtonBackground_Disable_Orange;
			COLORREF clr_RibbonButtonBackground_Disable_Yellow;
			COLORREF clr_RibbonButtonBackground_Disable_Green;
			COLORREF clr_RibbonButtonBackground_Disable_Blue;
			COLORREF clr_RibbonButtonBackground_Disable_Indigo;
			COLORREF clr_RibbonButtonBackground_Disable_Violet;

			COLORREF clr_RibbonButtonBackground_Pressed;
			COLORREF clr_RibbonButtonBackground_Pressed_Red;
			COLORREF clr_RibbonButtonBackground_Pressed_Orange;
			COLORREF clr_RibbonButtonBackground_Pressed_Yellow;
			COLORREF clr_RibbonButtonBackground_Pressed_Green;
			COLORREF clr_RibbonButtonBackground_Pressed_Blue;
			COLORREF clr_RibbonButtonBackground_Pressed_Indigo;
			COLORREF clr_RibbonButtonBackground_Pressed_Violet;

			COLORREF clr_RibbonButtonBackground_Highlight;
			COLORREF clr_RibbonButtonBackground_Highlight_Red;
			COLORREF clr_RibbonButtonBackground_Highlight_Orange;
			COLORREF clr_RibbonButtonBackground_Highlight_Yellow;
			COLORREF clr_RibbonButtonBackground_Highlight_Green;
			COLORREF clr_RibbonButtonBackground_Highlight_Blue;
			COLORREF clr_RibbonButtonBackground_Highlight_Indigo;
			COLORREF clr_RibbonButtonBackground_Highlight_Violet;

			COLORREF clr_RibbonButtonBorder;
			COLORREF clr_RibbonButtonBorder_Red;
			COLORREF clr_RibbonButtonBorder_Orange;
			COLORREF clr_RibbonButtonBorder_Yellow;
			COLORREF clr_RibbonButtonBorder_Green;
			COLORREF clr_RibbonButtonBorder_Blue;
			COLORREF clr_RibbonButtonBorder_Indigo;
			COLORREF clr_RibbonButtonBorder_Violet;

			COLORREF clr_RibbonButtonBorder_Disable;
			COLORREF clr_RibbonButtonBorder_Disable_Red;
			COLORREF clr_RibbonButtonBorder_Disable_Orange;
			COLORREF clr_RibbonButtonBorder_Disable_Yellow;
			COLORREF clr_RibbonButtonBorder_Disable_Green;
			COLORREF clr_RibbonButtonBorder_Disable_Blue;
			COLORREF clr_RibbonButtonBorder_Disable_Indigo;
			COLORREF clr_RibbonButtonBorder_Disable_Violet;

			COLORREF clr_RibbonButtonBorder_Pressed;
			COLORREF clr_RibbonButtonBorder_Pressed_Red;
			COLORREF clr_RibbonButtonBorder_Pressed_Orange;
			COLORREF clr_RibbonButtonBorder_Pressed_Yellow;
			COLORREF clr_RibbonButtonBorder_Pressed_Green;
			COLORREF clr_RibbonButtonBorder_Pressed_Blue;
			COLORREF clr_RibbonButtonBorder_Pressed_Indigo;
			COLORREF clr_RibbonButtonBorder_Pressed_Violet;

			COLORREF clr_RibbonButtonBorder_Highlight;
			COLORREF clr_RibbonButtonBorder_Highlight_Red;
			COLORREF clr_RibbonButtonBorder_Highlight_Orange;
			COLORREF clr_RibbonButtonBorder_Highlight_Yellow;
			COLORREF clr_RibbonButtonBorder_Highlight_Green;
			COLORREF clr_RibbonButtonBorder_Highlight_Blue;
			COLORREF clr_RibbonButtonBorder_Highlight_Indigo;
			COLORREF clr_RibbonButtonBorder_Highlight_Violet;

			COLORREF clr_RibbonButtonText;
			COLORREF clr_RibbonButtonText_Red;
			COLORREF clr_RibbonButtonText_Orange;
			COLORREF clr_RibbonButtonText_Yellow;
			COLORREF clr_RibbonButtonText_Green;
			COLORREF clr_RibbonButtonText_Blue;
			COLORREF clr_RibbonButtonText_Indigo;
			COLORREF clr_RibbonButtonText_Violet;

			COLORREF clr_RibbonButtonText_Disable;
			COLORREF clr_RibbonButtonText_Disable_Red;
			COLORREF clr_RibbonButtonText_Disable_Orange;
			COLORREF clr_RibbonButtonText_Disable_Yellow;
			COLORREF clr_RibbonButtonText_Disable_Green;
			COLORREF clr_RibbonButtonText_Disable_Blue;
			COLORREF clr_RibbonButtonText_Disable_Indigo;
			COLORREF clr_RibbonButtonText_Disable_Violet;

			COLORREF clr_RibbonButtonText_Pressed;
			COLORREF clr_RibbonButtonText_Pressed_Red;
			COLORREF clr_RibbonButtonText_Pressed_Orange;
			COLORREF clr_RibbonButtonText_Pressed_Yellow;
			COLORREF clr_RibbonButtonText_Pressed_Green;
			COLORREF clr_RibbonButtonText_Pressed_Blue;
			COLORREF clr_RibbonButtonText_Pressed_Indigo;
			COLORREF clr_RibbonButtonText_Pressed_Violet;

			COLORREF clr_RibbonButtonText_Highlight;
			COLORREF clr_RibbonButtonText_Highlight_Red;
			COLORREF clr_RibbonButtonText_Highlight_Orange;
			COLORREF clr_RibbonButtonText_Highlight_Yellow;
			COLORREF clr_RibbonButtonText_Highlight_Green;
			COLORREF clr_RibbonButtonText_Highlight_Blue;
			COLORREF clr_RibbonButtonText_Highlight_Indigo;
			COLORREF clr_RibbonButtonText_Highlight_Violet;
		};

		struct ColorRibbonSystemAndQuickButton
		{
			COLORREF clr_RibbonQuick_Background;
			
			COLORREF clr_RibbonQuickButton_Background;
			COLORREF clr_RibbonQuickButton_Background_Disable;
			COLORREF clr_RibbonQuickButton_Background_Highlight;
			COLORREF clr_RibbonQuickButton_Background_Pressed;
			
			COLORREF clr_RibbonQuickButton_Border;
			COLORREF clr_RibbonQuickButton_Border_Disable;
			COLORREF clr_RibbonQuickButton_Border_Highlight;
			COLORREF clr_RibbonQuickButton_Border_Pressed;

			COLORREF clr_RibbonSystemButton_Background;
			COLORREF clr_RibbonSystemButton_Background_Disable;
			COLORREF clr_RibbonSystemButton_Background_Highlight;
			COLORREF clr_RibbonSystemButton_Background_Pressed;
			
			COLORREF clr_RibbonSystemButton_Border;
			COLORREF clr_RibbonSystemButton_Border_Disable;
			COLORREF clr_RibbonSystemButton_Border_Highlight;
			COLORREF clr_RibbonSystemButton_Border_Pressed;

			COLORREF clr_RibbonSystemButton_Text;
			COLORREF clr_RibbonSystemButton_Text_Disable;
			COLORREF clr_RibbonSystemButton_Text_Highlight;
			COLORREF clr_RibbonSystemButton_Text_Pressed;

			COLORREF clr_RibbonSystemButton_TextQuickButton;
			COLORREF clr_RibbonSystemButton_TextQuickButton_Disable;
			COLORREF clr_RibbonSystemButton_TextQuickButton_Highlight;
			COLORREF clr_RibbonSystemButton_TextQuickButton_Pressed;

		};

		struct ColorStatusBarPane
		{
			COLORREF clr_StatusBarPane_Background;
			
			COLORREF clr_StatusBarPane_Background_Highlight;
			COLORREF clr_StatusBarPane_Background_Disable;

			COLORREF clr_StatusBarPane_ProgressBar;
			COLORREF clr_StatusBarPane_BarDest;
		
			COLORREF clr_StatusBarPane_Text;
			COLORREF clr_StatusBarPane_Text_Highlight;
			COLORREF clr_StatusBarPane_Text_Disable;
			COLORREF clr_StatusBarPane_TextProgressBar;
		};

		struct ColorControlElements
		{
			COLORREF clr_RichEdit_Background;
			COLORREF clr_RichEdit_Background_Highlight;
			COLORREF clr_RichEdit_Background_Disable;

			COLORREF clr_ProgressBar_Background;
			COLORREF clr_ProgressBar_Background_Red;
			COLORREF clr_ProgressBar_Background_Orange;
			COLORREF clr_ProgressBar_Background_Yellow;
			COLORREF clr_ProgressBar_Background_Green;
			COLORREF clr_ProgressBar_Background_Blue;
			COLORREF clr_ProgressBar_Background_Indigo;
			COLORREF clr_ProgressBar_Background_Violet;

			COLORREF clr_ProgressChunk_Background;
			COLORREF clr_ProgressChunk_Background_Red;
			COLORREF clr_ProgressChunk_Background_Orange;
			COLORREF clr_ProgressChunk_Background_Yellow;
			COLORREF clr_ProgressChunk_Background_Green;
			COLORREF clr_ProgressChunk_Background_Blue;
			COLORREF clr_ProgressChunk_Background_Indigo;
			COLORREF clr_ProgressChunk_Background_Violet;

			COLORREF clr_TextHiperLink;
			COLORREF clr_TextHiperLink_Red;
			COLORREF clr_TextHiperLink_Orange;
			COLORREF clr_TextHiperLink_Yellow;
			COLORREF clr_TextHiperLink_Green;
			COLORREF clr_TextHiperLink_Blue;
			COLORREF clr_TextHiperLink_Indigo;
			COLORREF clr_TextHiperLink_Violet;

			COLORREF clr_TextHiperLink_Highlight;
			COLORREF clr_TextHiperLink_Highlight_Red;
			COLORREF clr_TextHiperLink_Highlight_Orange;
			COLORREF clr_TextHiperLink_Highlight_Yellow;
			COLORREF clr_TextHiperLink_Highlight_Green;
			COLORREF clr_TextHiperLink_Highlight_Blue;
			COLORREF clr_TextHiperLink_Highlight_Indigo;
			COLORREF clr_TextHiperLink_Highlight_Violet;

		};

		struct ColorSliderZoom
		{
			COLORREF clr_SliderZoom_Background;
			COLORREF clr_SliderZoom_Background_Red;
			COLORREF clr_SliderZoom_Background_Orange;
			COLORREF clr_SliderZoom_Background_Yellow;
			COLORREF clr_SliderZoom_Background_Green;
			COLORREF clr_SliderZoom_Background_Blue;
			COLORREF clr_SliderZoom_Background_Indigo;
			COLORREF clr_SliderZoom_Background_Violet;

			COLORREF clr_BottonZoom_Background;
			COLORREF clr_BottonZoom_Background_Red;
			COLORREF clr_BottonZoom_Background_Orange;
			COLORREF clr_BottonZoom_Background_Yellow;
			COLORREF clr_BottonZoom_Background_Green;
			COLORREF clr_BottonZoom_Background_Blue;
			COLORREF clr_BottonZoom_Background_Indigo;
			COLORREF clr_BottonZoom_Background_Violet;

			COLORREF clr_BottonZoom_Background_Highlight;
			COLORREF clr_BottonZoom_Background_Highlight_Red;
			COLORREF clr_BottonZoom_Background_Highlight_Orange;
			COLORREF clr_BottonZoom_Background_Highlight_Yellow;
			COLORREF clr_BottonZoom_Background_Highlight_Green;
			COLORREF clr_BottonZoom_Background_Highlight_Blue;
			COLORREF clr_BottonZoom_Background_Highlight_Indigo;
			COLORREF clr_BottonZoom_Background_Highlight_Violet;

			COLORREF clr_BottonZoom_Background_Pressed;
			COLORREF clr_BottonZoom_Background_Pressed_Red;
			COLORREF clr_BottonZoom_Background_Pressed_Orange;
			COLORREF clr_BottonZoom_Background_Pressed_Yellow;
			COLORREF clr_BottonZoom_Background_Pressed_Green;
			COLORREF clr_BottonZoom_Background_Pressed_Blue;
			COLORREF clr_BottonZoom_Background_Pressed_Indigo;
			COLORREF clr_BottonZoom_Background_Pressed_Violet;

			COLORREF clr_BottonZoom_Background_Disabled;
			COLORREF clr_BottonZoom_Background_Disabled_Red;
			COLORREF clr_BottonZoom_Background_Disabled_Orange;
			COLORREF clr_BottonZoom_Background_Disabled_Yellow;
			COLORREF clr_BottonZoom_Background_Disabled_Green;
			COLORREF clr_BottonZoom_Background_Disabled_Blue;
			COLORREF clr_BottonZoom_Background_Disabled_Indigo;
			COLORREF clr_BottonZoom_Background_Disabled_Violet;

			COLORREF clr_BottonZoom_Background_Plus;
			COLORREF clr_BottonZoom_Background_Plus_Red;
			COLORREF clr_BottonZoom_Background_Plus_Orange;
			COLORREF clr_BottonZoom_Background_Plus_Yellow;
			COLORREF clr_BottonZoom_Background_Plus_Green;
			COLORREF clr_BottonZoom_Background_Plus_Blue;
			COLORREF clr_BottonZoom_Background_Plus_Indigo;
			COLORREF clr_BottonZoom_Background_Plus_Violet;

			COLORREF clr_BottonZoom_Background_Minus;
			COLORREF clr_BottonZoom_Background_Minus_Red;
			COLORREF clr_BottonZoom_Background_Minus_Orange;
			COLORREF clr_BottonZoom_Background_Minus_Yellow;
			COLORREF clr_BottonZoom_Background_Minus_Green;
			COLORREF clr_BottonZoom_Background_Minus_Blue;
			COLORREF clr_BottonZoom_Background_Minus_Indigo;
			COLORREF clr_BottonZoom_Background_Minus_Violet;

			COLORREF clr_BottonZoom_Background_Plus_Disabled;
			COLORREF clr_BottonZoom_Background_Plus_Disabled_Red;
			COLORREF clr_BottonZoom_Background_Plus_Disabled_Orange;
			COLORREF clr_BottonZoom_Background_Plus_Disabled_Yellow;
			COLORREF clr_BottonZoom_Background_Plus_Disabled_Green;
			COLORREF clr_BottonZoom_Background_Plus_Disabled_Blue;
			COLORREF clr_BottonZoom_Background_Plus_Disabled_Indigo;
			COLORREF clr_BottonZoom_Background_Plus_Disabled_Violet;

			COLORREF clr_BottonZoom_Background_Minus_Disabled;
			COLORREF clr_BottonZoom_Background_Minus_Disabled_Red;
			COLORREF clr_BottonZoom_Background_Minus_Disabled_Orange;
			COLORREF clr_BottonZoom_Background_Minus_Disabled_Yellow;
			COLORREF clr_BottonZoom_Background_Minus_Disabled_Green;
			COLORREF clr_BottonZoom_Background_Minus_Disabled_Blue;
			COLORREF clr_BottonZoom_Background_Minus_Disabled_Indigo;
			COLORREF clr_BottonZoom_Background_Minus_Disabled_Violet;

			COLORREF clr_BottonZoom_Background_Plus_Highlight;
			COLORREF clr_BottonZoom_Background_Plus_Highlight_Red;
			COLORREF clr_BottonZoom_Background_Plus_Highlight_Orange;
			COLORREF clr_BottonZoom_Background_Plus_Highlight_Yellow;
			COLORREF clr_BottonZoom_Background_Plus_Highlight_Green;
			COLORREF clr_BottonZoom_Background_Plus_Highlight_Blue;
			COLORREF clr_BottonZoom_Background_Plus_Highlight_Indigo;
			COLORREF clr_BottonZoom_Background_Plus_Highlight_Violet;

			COLORREF clr_BottonZoom_Background_Minus_Highlight;
			COLORREF clr_BottonZoom_Background_Minus_Highlight_Red;
			COLORREF clr_BottonZoom_Background_Minus_Highlight_Orange;
			COLORREF clr_BottonZoom_Background_Minus_Highlight_Yellow;
			COLORREF clr_BottonZoom_Background_Minus_Highlight_Green;
			COLORREF clr_BottonZoom_Background_Minus_Highlight_Blue;
			COLORREF clr_BottonZoom_Background_Minus_Highlight_Indigo;
			COLORREF clr_BottonZoom_Background_Minus_Highlight_Violet;

			COLORREF clr_BottonZoom_Background_Plus_Pressed;
			COLORREF clr_BottonZoom_Background_Plus_Pressed_Red;
			COLORREF clr_BottonZoom_Background_Plus_Pressed_Orange;
			COLORREF clr_BottonZoom_Background_Plus_Pressed_Yellow;
			COLORREF clr_BottonZoom_Background_Plus_Pressed_Green;
			COLORREF clr_BottonZoom_Background_Plus_Pressed_Blue;
			COLORREF clr_BottonZoom_Background_Plus_Pressed_Indigo;
			COLORREF clr_BottonZoom_Background_Plus_Pressed_Violet;

			COLORREF clr_BottonZoom_Background_Minus_Pressed;
			COLORREF clr_BottonZoom_Background_Minus_Pressed_Red;
			COLORREF clr_BottonZoom_Background_Minus_Pressed_Orange;
			COLORREF clr_BottonZoom_Background_Minus_Pressed_Yellow;
			COLORREF clr_BottonZoom_Background_Minus_Pressed_Green;
			COLORREF clr_BottonZoom_Background_Minus_Pressed_Blue;
			COLORREF clr_BottonZoom_Background_Minus_Pressed_Indigo;
			COLORREF clr_BottonZoom_Background_Minus_Pressed_Violet;

		};

		// ----------------------------------------------------------------------------
		// ПАРАМЕТРЫ ОТРИСОВКИ / PARAMS DRAWING
		// ----------------------------------------------------------------------------
		struct ParametrsCMenuImage
		{
			CMenuImages::IMAGE_STATE img_ColorImage;
			COLORREF clr_UserColorImageIcon;
		};

		struct ParametrsDrawStandard
		{
			BOOL m_bStandartDrawingTab = FALSE; //произовдить отрисовку OnDrawTab или другими методами

			BOOL m_bUnderlineTabs_AutoHideTab = TRUE;
			BOOL m_bUnderlineTabs_MDITab = TRUE;

			BOOL m_bBorderUnderLine_AutoHideTab = FALSE;
			BOOL m_bBorderUnderLine_MDITab = FALSE;

			int m_iSizeUnderlineTabs_AutoHideTab = 10;
			int m_iSizeUnderlineTabs_MDITab = 10;

			int m_iSizeBorder_CaptionBar_Top = 0;
			int m_iSizeBorder_CaptionBar_Bottom = 0;
			int m_iSizeBorder_CaptionBar_Right = 3;
			int m_iSizeBorder_CaptionBar_Left = 3;

			int m_iSizeBorder_AutoHideTab_Top = 0;
			int m_iSizeBorder_AutoHideTab_Bottom = 3;
			int m_iSizeBorder_AutoHideTab_Right = 0;
			int m_iSizeBorder_AutoHideTab_Left = 0;

			int m_iSizeBorder_MDITab_Top = 0;
			int m_iSizeBorder_MDITab_Bottom = 0;
			int m_iSizeBorder_MDITab_Right = 0;
			int m_iSizeBorder_MDITab_Left = 0;

			int m_iSizeBorder_MDIButtonTab_Top = 0;
			int m_iSizeBorder_MDIButtonTab_Bottom = 2;
			int m_iSizeBorder_MDIButtonTab_Right = 2;
			int m_iSizeBorder_MDIButtonTab_Left = 0;

		};

		struct ParametrsDrawRibbon
		{
			/*TabsRibbon*/
			BOOL m_bParamRibbon_DrawTabsBackground_BorderAllTabsCaption = TRUE; //рисовать границу всех вкладок
			BOOL m_bParamRibbon_DrawTabsBackground_BorderNoneTabs = FALSE; //рисовать границу только внеконтекстных
			BOOL m_bParamRibbon_DrawTabsBackground_BorderColorTabs = FALSE; //рисовать границу только контекстных

			/*PanelRibbon*/

			/*CategoryRibbon*/

			int m_iSizeBorder_ContextTop = 0;
			int m_iSizeBorder_ContextBottom = 0;
			int m_iSizeBorder_ContextRight = 0;
			int m_iSizeBorder_ContextLeft = 0;

			int m_iSizeBorder_TabsTop = 0;
			int m_iSizeBorder_TabsBottom = 0;
			int m_iSizeBorder_TabsRight = 0;
			int m_iSizeBorder_TabsLeft = 0;

			int m_iSizeBorder_CategoryTop = 2;
			int m_iSizeBorder_CategoryBottom = 0;
			int m_iSizeBorder_CategoryRight = 2;
			int m_iSizeBorder_CategoryLeft = 0;

			int m_iSizeBorder_PanelTop = 0;
			int m_iSizeBorder_PanelBottom = 0;
			int m_iSizeBorder_PanelRight = 0;
			int m_iSizeBorder_PanelLeft = 0;

			int m_iSizeBorder_PanelLabelTop = 0;
			int m_iSizeBorder_PanelLabelBottom = 0;
			int m_iSizeBorder_PanelLabelRight = 0;
			int m_iSizeBorder_PanelLabelLeft = 0;

			int m_iSizeBorder_RibbonButtonTop = 0;
			int m_iSizeBorder_RibbonButtonBottom = 0;
			int m_iSizeBorder_RibbonButtonRight = 0;
			int m_iSizeBorder_RibbonButtonLeft = 0;
		};


		/*Standard*/
		ColorGlobalElement m_ColorGlobalElement;
		ColorMiniFrame m_ColorMiniFrame;
		ColorTabs m_ColorTabs;
		ColorPropertyGrid m_ColorPropertyGrid;
		ColorThreeView m_ColorThreeView;
		ColorToolbars m_ColorToolbars;
		ColorPaneCaptionsAndButton m_ColorPaneCaptionsAndButton;
		ColorOutlookBar m_ColorOutlookBar;

		/*Ribbon*/
		ColorRibbonBackground m_ColorRibbonBackground;
		ColorMainMenuRibbon m_ColorMainMenuRibbon;
		ColorRibbonTabs m_ColorRibbonTabs;
		ColorRibbonCategory m_ColorRibbonCategory;
		ColorRibbonPanel m_ColorRibbonPanel;
		ColorRibbonButton m_ColorRibbonButton;
		ColorStatusBarPane m_ColorStatusBarPane;
		ColorControlElements m_ColorContrlElem;
		ColorSliderZoom m_ColorSliderZoom;
		ColorRibbonSystemAndQuickButton m_ColorRibbonSystemAndQuickButton;

		/*Params*/
		ParametrsDrawStandard m_sParamStandard;
		ParametrsDrawRibbon m_sParamRibbon;
		ParametrsCMenuImage m_sParamCMenu;

	};

	enum ThemeStyle {
		nopStyle_None,
		/*---*/
		nopStyle_DarkGray,
		nopStyle_White,
		nopStyle_Blue,
		nopStyle_Green,
		nopStyle_Purple,
		/*---*/
		nopStyle_DarkBlue,
		nopStyle_DarkGreen,
		nopStyle_DarkYollow,
		nopStyle_DarkPurple,
		/*---*/
		nopStyle_WhiteBlue,
		nopStyle_WhiteGreen,
		nopStyle_WhiteYollow,
		nopStyle_WhitePurple,
		/*---*/
		nopStyle_SystemColorLight,
		nopStyle_SystemColorDark,
		/*---*/
		nopStyle_UserStyle
	};

private:

protected:

	
	static ThemeStyle m_currentStyle;
	static ColorTheme m_currentTheme;
	static DefaultColorTheme m_DefaultColor;

	static CMenuImages m_CMIUserImage;

	CNPUP_VisualManager();
	~CNPUP_VisualManager();

	void NVS_DrawAlphaChanel(CDC* clientDC, BYTE alphaChanel);
	void NVS_DrawRectBorder(CDC* clientDC, CRect rectDraw, COLORREF colorRGB,
		int sizeTopBorder, int sizeBottomBorder, int sizeLeftBorder, int sizeRightBorder);
	
	static void SetColorThemeGlobal();
	static void SetChildWindowStyle();

public:

	static void SetStyle(ThemeStyle style);
	ThemeStyle GetStyle() { return m_currentStyle; }

	virtual void OnUpdateSystemColors();

public:

	// ============================================================================
	// СИСТЕМНЫЕ ФУНКЦИИ / SYSTEM FUNCTION
	// ----------------------------------------------------------------------------
	// Главный фрейм, неклиентские области
	// Frame and Non-client area
	// ============================================================================
#pragma region SystemFunction

	virtual BOOL IsOwnerDrawCaption() { return FALSE; } // управляет включением работы по отрисовке неклиенсткой области средствами менеджера
	virtual BOOL OnNcPaint(CWnd* pWnd, const CObList& lstSysButtons, CRect rectRedraw);
	virtual BOOL OnNcActivate(CWnd* pWnd, BOOL bActive);


#pragma endregion

	// ============================================================================
	// КЛАССИЧЕСКИЕ ЭЛЕМЕНТЫ ИНТЕРФЕЙСА / STANDARD UI FUNCTIONS
	// ----------------------------------------------------------------------------
	// Фрейм и мини-фреймы, прочие классические элементы
	// Frame and Mini-frame, other standard element
	// ============================================================================
#pragma region Standard

	// MainFrame
	virtual void OnFillBarBackground(CDC* pDC, CBasePane* pBar, CRect rectClient, CRect rectClip, BOOL bNCArea = FALSE);
	virtual COLORREF GetToolbarButtonTextColor(CMFCToolBarButton* pButton, CMFCVisualManager::AFX_BUTTON_STATE state);
	virtual COLORREF GetMenuItemTextColor(CMFCToolBarMenuButton* pButton, BOOL bHighlighted, BOOL bDisabled);

	virtual COLORREF OnDrawPaneCaption(CDC* pDC, CDockablePane* pBar, BOOL bActive, CRect rectCaption, CRect rectButtons);
	virtual COLORREF OnFillCaptionBarButton(CDC* pDC, CMFCCaptionBar* pBar, CRect rect,
		BOOL bIsPressed, BOOL bIsHighlighted, BOOL bIsDisabled, BOOL bHasDropDownArrow, BOOL bIsSysButton);
	virtual void OnDrawCaptionBarInfoArea(CDC* pDC, CMFCCaptionBar* pBar, CRect rect);
	virtual COLORREF GetCaptionBarTextColor(CMFCCaptionBar* pBar);
	virtual void OnDrawCaptionBarButtonBorder(CDC* pDC, CMFCCaptionBar* pBar, CRect rect, BOOL bIsPressed,
		BOOL bIsHighlighted, BOOL bIsDisabled, BOOL bHasDropDownArrow, BOOL bIsSysButton);

	virtual void OnDrawShowAllMenuItems(CDC* pDC, CRect rect, CMFCVisualManager::AFX_BUTTON_STATE state);
	virtual void OnDrawMenuResizeBar(CDC* pDC, CRect rect, int nResizeFlags);

	virtual void OnDrawStatusBarPaneBorder(CDC* pDC, CMFCStatusBar* pBar, CRect rectPane, UINT uiID, UINT nStyle);
	virtual void OnDrawStatusBarSizeBox(CDC* pDC, CMFCStatusBar* pStatBar, CRect rectSizeBox);
	virtual COLORREF GetStatusBarPaneTextColor(CMFCStatusBar* pStatusBar, CMFCStatusBarPaneInfo* pPane);

	virtual void OnDrawMenuBorder(CDC* pDC, CMFCPopupMenu* pMenu, CRect rect);
	virtual void OnDrawMenuShadow(CDC* pDC, const CRect& rectClient, const CRect& rectExclude, int nDepth,
		int iMinBrightness, int iMaxBrightness, CBitmap* pBmpSaveBottom, CBitmap* pBmpSaveRight, BOOL bRTL);
	virtual void OnDrawBarGripper(CDC* pDC, CRect rectGripper, BOOL bHorz, CBasePane* pBar);
	virtual void OnDrawSeparator(CDC* pDC, CBasePane* pBar, CRect rect, BOOL bIsHoriz);
	virtual COLORREF OnDrawMenuLabel(CDC* pDC, CRect rect);
	virtual void OnDrawComboDropButton(CDC* pDC, CRect rect, BOOL bDisabled, BOOL bIsDropped, BOOL bIsHighlighted, CMFCToolBarComboBoxButton* pButton);
	virtual BOOL OnEraseMDIClientArea(CDC* pDC, CRect rectClient);
	

	// Auto-hide buttons:
	virtual void OnFillAutoHideButtonBackground(CDC* pDC, CRect rect, CMFCAutoHideButton* pButton);
	virtual void OnDrawAutoHideButtonBorder(CDC* pDC, CRect rectBounds, CRect rectBorderSize, CMFCAutoHideButton* pButton);
	virtual COLORREF GetAutoHideButtonTextColor(CMFCAutoHideButton* pButton);

	// Mini-frame:
	virtual COLORREF OnFillMiniFrameCaption(CDC* pDC, CRect rectCaption, CPaneFrameWnd* pFrameWnd, BOOL bActive);
	virtual void OnDrawMiniFrameBorder(CDC* pDC, CPaneFrameWnd* pFrameWnd, CRect rectBorder, CRect rectBorderSize);
	virtual void OnDrawCaptionButton(CDC* pDC, CMFCCaptionButton* pButton, BOOL bActive, BOOL bHorz, BOOL bMaximized, BOOL bDisabled, int nImageID = -1);
	virtual void OnDrawPaneDivider(CDC* pDC, CPaneDivider* pSlider, CRect rect, BOOL bAutoHideMode);

	// Property Grid
	virtual COLORREF GetPropertyGridGroupColor(CMFCPropertyGridCtrl* pPropList);
	virtual COLORREF GetPropertyGridGroupTextColor(CMFCPropertyGridCtrl* pPropList);
	virtual void OnDrawExpandingBox(CDC* pDC, CRect rect, BOOL bIsOpened, COLORREF colorBox);

	// Tabs
	virtual void GetTabFrameColors(const CMFCBaseTabCtrl* pTabWnd, COLORREF& clrDark, COLORREF& clrBlack,
		COLORREF& clrHighlight, COLORREF& clrFace, COLORREF& clrDarkShadow, COLORREF& clrLight, CBrush*& pbrFace, CBrush*& pbrBlack);
	virtual void OnEraseTabsArea(CDC* pDC, CRect rect, const CMFCBaseTabCtrl* pTabWnd);
	virtual BOOL OnEraseTabsFrame(CDC* pDC, CRect rect, const CMFCBaseTabCtrl* pTabWnd);
	virtual void OnEraseTabsButton(CDC* pDC, CRect rect, CMFCButton* pButton, CMFCBaseTabCtrl* pWndTab);
	virtual void OnDrawTabsButtonBorder(CDC* pDC, CRect& rect, CMFCButton* pButton, UINT uiState, CMFCBaseTabCtrl* pWndTab);
	virtual void OnDrawTab(CDC* pDC, CRect rectTab, int iTab, BOOL bIsActive, const CMFCBaseTabCtrl* pTabWnd);
	virtual void OnDrawTabCloseButton(CDC* pDC, CRect rect, const CMFCBaseTabCtrl* pTabWnd, BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsDisabled);
	virtual void OnDrawTabContent(CDC* pDC, CRect rectTab, int iTab, BOOL bIsActive, const CMFCBaseTabCtrl* pTabWnd, COLORREF clrText);
	virtual void OnDrawTabResizeBar(CDC* pDC, CMFCBaseTabCtrl* pWndTab, BOOL bIsVert, CRect rect, CBrush* pbrFace, CPen* pPen);
	virtual void OnFillTab(CDC* pDC, CRect rectFill, CBrush* pbrFill, int iTab, BOOL bIsActive, const CMFCBaseTabCtrl* pTabWnd);

	// Outlook bar
	virtual void OnFillOutlookPageButton(CDC* pDC, const CRect& rect, BOOL bIsHighlighted, BOOL bIsPressed, COLORREF& clrText);
	virtual void OnDrawOutlookPageButtonBorder(CDC* pDC, CRect& rectBtn, BOOL bIsHighlighted, BOOL bIsPressed);
	virtual void OnDrawOutlookBarSplitter(CDC* pDC, CRect rectSplitter);
	virtual void OnFillOutlookBarCaption(CDC* pDC, CRect rectCaption, COLORREF& clrText);

	// Toolbar and menu
	virtual void OnDrawFloatingToolbarBorder(CDC* pDC, CMFCBaseToolBar* pToolBar, CRect rectBorder, CRect rectBorderSize);
	virtual void OnDrawButtonBorder(CDC* pDC, CMFCToolBarButton* pButton, CRect rect, CMFCVisualManager::AFX_BUTTON_STATE state);
	virtual void OnHighlightRarelyUsedMenuItems(CDC* pDC, CRect rectRarelyUsed);
	virtual void OnHighlightMenuItem(CDC *pDC, CMFCToolBarMenuButton* pButton, CRect rect, COLORREF& clrText);
	virtual void OnDrawMenuCheck(CDC* pDC, CMFCToolBarMenuButton* pButton, CRect rect, BOOL bHighlight, BOOL bIsRadio);
	virtual void OnFillButtonInterior(CDC* pDC, CMFCToolBarButton* pButton, CRect rect, CMFCVisualManager::AFX_BUTTON_STATE state);
	
	virtual BOOL IsToolbarRoundShape(CMFCToolBar* pToolBar);

	// Other
	virtual COLORREF OnFillCommandsListBackground(CDC* pDC, CRect rect, BOOL bIsSelected = FALSE);
	virtual void OnDrawCheckBoxEx(CDC *pDC, CRect rect, int nState, BOOL bHighlighted, BOOL bPressed, BOOL bEnabled);
	virtual void OnDrawSpinButtons(CDC* pDC, CRect rectSpin, int nState, BOOL bOrientation, CMFCSpinButtonCtrl* pSpinCtrl);
	virtual BOOL GetToolTipInfo(CMFCToolTipInfo& params, UINT nType = (UINT)(-1));
	virtual void OnFillMenuImageRect(CDC* pDC, CMFCToolBarButton* pButton, CRect rect, CMFCVisualManager::AFX_BUTTON_STATE state);


#pragma endregion

	// ============================================================================
	// ЭЛЕМЕНТЫ ЛЕНТОЧНОГО ИНТЕРФЕЙСА / RIBBON UI FUNCTIONS
	// ----------------------------------------------------------------------------
	// Методы отрисовки компонентов Ribbon-панелей
	// Rendering methods for Ribbon interface components
	// ============================================================================
#pragma region Ribbon

	// Main frame
	virtual void OnDrawRibbonCaption(CDC* pDC, CMFCRibbonBar* pBar, CRect rectCaption, CRect rectText);
	virtual COLORREF OnDrawRibbonCategoryCaption(CDC* pDC, CMFCRibbonContextCaption* pContextCaption);
	virtual COLORREF OnDrawRibbonStatusBarPane(CDC* pDC, CMFCRibbonStatusBar* pBar, CMFCRibbonStatusBarPane* pPane);

	// Tabs
	virtual COLORREF OnDrawRibbonTabsFrame(CDC* pDC, CMFCRibbonBar* pWndRibbonBar, CRect rectTab);
	virtual void OnDrawRibbonCategory(CDC* pDC, CMFCRibbonCategory* pCategory, CRect rectCategory);
	virtual COLORREF OnDrawRibbonCategoryTab(CDC* pDC, CMFCRibbonTab* pTab, BOOL bIsActive);
	virtual COLORREF OnDrawRibbonPanel(CDC* pDC, CMFCRibbonPanel* pPanel, CRect rectPanel, CRect rectCaption);
	virtual void OnDrawRibbonPanelCaption(CDC* pDC, CMFCRibbonPanel* pPanel, CRect rectCaption);
	virtual COLORREF OnFillRibbonButton(CDC* pDC, CMFCRibbonButton* pButton);
	virtual void OnDrawRibbonButtonBorder(CDC* pDC, CMFCRibbonButton* pButton);

	// Main menu Ribbon
	virtual void OnDrawRibbonApplicationButton(CDC* pDC, CMFCRibbonButton* pButton);
	virtual void OnDrawRibbonMainPanelFrame(CDC* pDC, CMFCRibbonMainPanel* pPanel, CRect rect);
	virtual COLORREF OnFillRibbonMainPanelButton(CDC* pDC, CMFCRibbonButton* pButton);
	virtual void OnDrawRibbonMainPanelButtonBorder(CDC* pDC, CMFCRibbonButton* pButton); //--

	// Other
	virtual COLORREF GetRibbonEditBackgroundColor(CMFCRibbonRichEditCtrl* pEdit, BOOL bIsHighlighted, BOOL bIsPaneHighlighted, BOOL bIsDisabled);
	virtual void OnDrawRibbonMenuCheckFrame(CDC* pDC, CMFCRibbonButton* pButton, CRect rect);
	virtual void OnDrawRibbonSliderZoomButton(CDC* pDC, CMFCRibbonSlider* pSlider, CRect rect, BOOL bIsZoomOut, BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsDisabled);
	virtual void OnDrawRibbonSliderChannel(CDC* pDC, CMFCRibbonSlider* pSlider, CRect rect);
	virtual void OnDrawRibbonSliderThumb(CDC* pDC, CMFCRibbonSlider* pSlider, CRect rect, BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsDisabled);
	virtual void OnDrawRibbonProgressBar(CDC* pDC, CMFCRibbonProgressBar* pProgress, CRect rectProgress, CRect rectChunk, BOOL bInfiniteMode);
	virtual COLORREF GetRibbonHyperlinkTextColor(CMFCRibbonLinkCtrl* pHyperLink);
	virtual void OnDrawRibbonLabel(CDC* pDC, CMFCRibbonLabel* pLabel, CRect rect);

#pragma endregion

	// ============================================================================
	// МЕТОДЫ НЕ ВОШЕДШИЕ В РЕЛИЗ ПО РАЗНЫМ ПРИЧИНАМ
	// ----------------------------------------------------------------------------
	// METHODS NOT INCLUDED IN THE RELEASE FOR VARIOUS REASONS
	// ============================================================================
#pragma region NotIncluded

	//virtual int GetRibbonPopupBorderSize(const CMFCRibbonPanelMenu* pPopup) const;
	//virtual void OnFillRibbonMenuFrame(CDC* pDC, CMFCRibbonMainPanel* pPanel, CRect rect);
	//virtual void OnDrawRibbonRecentFilesFrame(CDC* pDC, CMFCRibbonMainPanel* pPanel, CRect rect);
	//virtual void OnDrawRibbonDefaultPaneButton(CDC* pDC, CMFCRibbonButton* pButton);
	//virtual void OnDrawRibbonGalleryButton(CDC* pDC, CMFCRibbonGalleryIcon* pButton); //этот метод кто-то перехватывает
	//virtual void OnDrawRibbonGalleryBorder(CDC* pDC, CMFCRibbonGallery* pButton, CRect rectBorder); //этот тоже
	//virtual void OnDrawRibbonColorPaletteBox(CDC* pDC, CMFCRibbonColorButton* pColorButton, CMFCRibbonGalleryIcon* pIcon,
	//	COLORREF color, CRect rect, BOOL bDrawTopEdge, BOOL bDrawBottomEdge, BOOL bIsHighlighted, BOOL bIsChecked, BOOL bIsDisabled);
	//virtual COLORREF GetTabTextColor(const CMFCBaseTabCtrl* pTabWnd, int iTab, BOOL bIsActive);
	//virtual int GetTabHorzMargin(const CMFCBaseTabCtrl* pTabWnd);
	//virtual void OnDrawStatusBarProgress(CDC* pDC, CMFCStatusBar* pStatusBar, CRect rectProgress, int nProgressTotal, int nProgressCurr,
	//	COLORREF clrBar, COLORREF clrProgressBarDest, COLORREF clrProgressText, BOOL bProgressText);
	//virtual int GetShowAllMenuItemsHeight(CDC* pDC, const CSize& sizeDefault);
	//virtual void GetSmartDockingBaseGuideColors(COLORREF& clrBaseGroupBackground, COLORREF& clrBaseGroupBorder);

	//virtual void OnDrawComboBorder(CDC* pDC, CRect rect, BOOL bDisabled, BOOL bIsDropped, BOOL bIsHighlighted, CMFCToolBarComboBoxButton* pButton);
	//virtual COLORREF OnDrawPropertySheetListItem(CDC* pDC, CMFCPropertySheet* pParent, CRect rect, BOOL bIsHighlihted, BOOL bIsSelected);
	//virtual void DrawCustomizeButton(CDC* pDC, CRect rect, BOOL bIsHorz, CMFCVisualManager::AFX_BUTTON_STATE state, BOOL bIsCustomize, BOOL bIsMoreButtons);
	//virtual void OnDrawCaptionButtonIcon(CDC* pDC, CMFCCaptionButton* pButton, CMenuImages::IMAGES_IDS id, BOOL bActive, BOOL bDisabled, CPoint ptImage);
	//virtual COLORREF GetToolbarDisabledTextColor();
	//virtual BOOL DrawStatusBarProgress(CDC* pDC, CMFCStatusBar* pStatusBar, CRect rectProgress, int nProgressTotal, int nProgressCurr,
	//	COLORREF clrBar, COLORREF clrProgressBarDest, COLORREF clrProgressText, BOOL bProgressText);

#pragma endregion

};


