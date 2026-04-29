// Этот исходный код примеров MFC демонстрирует функционирование пользовательского интерфейса Fluent на основе MFC в Microsoft Office
// ("Fluent UI") и предоставляется исключительно как справочный материал в качестве дополнения к
// справочнику по пакету Microsoft Foundation Classes и связанной электронной документации,
// включенной в программное обеспечение библиотеки MFC C++.
// Условия лицензионного соглашения на копирование, использование или распространение Fluent UI доступны отдельно.
// Для получения дополнительных сведений о нашей программе лицензирования Fluent UI посетите веб-сайт
// https://go.microsoft.com/fwlink/?LinkId=238214.
//
// (C) Корпорация Майкрософт (Microsoft Corp.)
// Все права защищены.

// MainFrm.cpp: реализация класса CMainFrame
//

#include "pch.h"
#include "framework.h"
#include "NOP_PUP_VSMNG.h"
#include "MainFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif
#pragma comment(lib, "dwmapi.lib")

// CMainFrame

IMPLEMENT_DYNAMIC(CMainFrame, CMDIFrameWndEx)

BEGIN_MESSAGE_MAP(CMainFrame, CMDIFrameWndEx)
	ON_WM_CREATE()
	ON_COMMAND(ID_WINDOW_MANAGER, &CMainFrame::OnWindowManager)
	ON_COMMAND_RANGE(ID_VIEW_APPLOOK_WINDOWS_7, ID_BUTTON_NOP_SYSCLR_DARK, &CMainFrame::OnApplicationLook)
	ON_UPDATE_COMMAND_UI_RANGE(ID_VIEW_APPLOOK_WIN_2000, ID_VIEW_APPLOOK_WINDOWS_7, &CMainFrame::OnUpdateApplicationLook)
	ON_COMMAND(ID_VIEW_CAPTION_BAR, &CMainFrame::OnViewCaptionBar)
	ON_UPDATE_COMMAND_UI(ID_VIEW_CAPTION_BAR, &CMainFrame::OnUpdateViewCaptionBar)
	ON_COMMAND(ID_TOOLS_OPTIONS, &CMainFrame::OnOptions)
	ON_WM_SETTINGCHANGE()
END_MESSAGE_MAP()

// Создание или уничтожение CMainFrame

CMainFrame::CMainFrame() noexcept
{
	// TODO: добавьте код инициализации члена
	theApp.m_nAppLook = theApp.GetInt(_T("ApplicationLook"), ID_VIEW_APPLOOK_VS_2008);
}

CMainFrame::~CMainFrame()
{

}



int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CMDIFrameWndEx::OnCreate(lpCreateStruct) == -1)
		return -1;

	BOOL bNameValid;

	//CMDITabInfo mdiTabParams;
	//mdiTabParams.m_style = CMFCTabCtrl::STYLE_3D; // другие доступные стили...
	//mdiTabParams.m_bActiveTabCloseButton = FALSE;      // установите значение FALSE, чтобы расположить кнопку \"Закрыть\" в правой части области вкладки
	//mdiTabParams.m_bTabIcons = FALSE;    // установите значение TRUE, чтобы включить значки документов на вкладках MDI
	//mdiTabParams.m_bAutoColor = TRUE;    // установите значение FALSE, чтобы отключить автоматическое выделение цветом вкладок MDI
	//mdiTabParams.m_bDocumentMenu = TRUE; // включить меню документа на правой границе области вкладки
	//EnableMDITabbedGroups(TRUE, mdiTabParams);

	CMDITabInfo mdiTabParams;
	mdiTabParams.m_style = CMFCTabCtrl::STYLE_3D_SCROLLED; // другие доступные стили...
	mdiTabParams.m_bActiveTabCloseButton = TRUE;      // установите значение FALSE, чтобы расположить кнопку \"Закрыть\" в правой части области вкладки
	mdiTabParams.m_bTabIcons = FALSE;    // установите значение TRUE, чтобы включить значки документов на вкладках MDI
	mdiTabParams.m_bAutoColor = FALSE;    // установите значение FALSE, чтобы отключить автоматическое выделение цветом вкладок MDI
	mdiTabParams.m_bDocumentMenu = TRUE; // включить меню документа на правой границе области вкладки
	mdiTabParams.m_nTabBorderSize = 0;
	mdiTabParams.m_bTabCustomTooltips = TRUE;
	mdiTabParams.m_bFlatFrame = TRUE;
	mdiTabParams.m_bReuseRemovedTabGroups = TRUE;
	mdiTabParams.m_bTabCloseButton = TRUE;
	EnableMDITabbedGroups(TRUE, mdiTabParams);


	CMFCTabCtrl &currentControl = m_wndClientArea.GetMDITabs();
	




	m_wndRibbonBar.Create(this);
	m_wndRibbonBar.LoadFromResource(IDR_RIBBON);
	m_wndRibbonBar.ShowContextCategories(ID_CONTEXT3, TRUE);
	m_wndRibbonBar.ShowContextCategories(ID_CONTEXT2, TRUE);
	m_wndRibbonBar.ShowContextCategories(ID_CONTEXT1, TRUE);

	if (!m_wndStatusBar.Create(this))
	{
		TRACE0("Не удалось создать строку состояния\n");
		return -1;      // не удалось создать
	}

	CString strTitlePane1;
	CString strTitlePane2;
	bNameValid = strTitlePane1.LoadString(IDS_STATUS_PANE1);
	ASSERT(bNameValid);
	bNameValid = strTitlePane2.LoadString(IDS_STATUS_PANE2);
	ASSERT(bNameValid);
	m_wndStatusBar.AddElement(new CMFCRibbonStatusBarPane(ID_STATUSBAR_PANE1, strTitlePane1, TRUE), strTitlePane1);
	m_wndStatusBar.AddExtendedElement(new CMFCRibbonStatusBarPane(ID_STATUSBAR_PANE2, strTitlePane2, TRUE), strTitlePane2);
	


	// включить режим работы закрепляемых окон стилей Visual Studio 2005
	//CDockingManager::SetDockingMode(DT_SMART, AFX_SDT_VS2005);
	// включить режим работы автоматического скрытия закрепляемых окон стилей Visual Studio 2005
	EnableAutoHidePanes(CBRS_ALIGN_ANY);

	// Создать заголовок окна:
	if (!CreateCaptionBar())
	{
		TRACE0("Не удалось создать заголовок окна\n");
		return -1;      // не удалось создать
	}

	// Загрузить изображение элемента меню (не размещенное на каких-либо стандартных панелях инструментов):
	CMFCToolBar::AddToolBarForImageCollection(IDR_MENU_IMAGES, theApp.m_bHiColorIcons ? IDB_MENU_IMAGES_24 : 0);

	// создать закрепляемые окна
	if (!CreateDockingWindows())
	{
		TRACE0("Не удалось создать закрепляемые окна\n");
		return -1;
	}

	m_wndFileView.EnableDocking(CBRS_ALIGN_ANY);
	m_wndClassView.EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_wndFileView);
	CDockablePane* pTabbedBar = nullptr;
	m_wndClassView.AttachToTabWnd(&m_wndFileView, DM_SHOW, TRUE, &pTabbedBar);
	m_wndOutput.EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_wndOutput);
	m_wndProperties.EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_wndProperties);


	m_TaskPane.EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_TaskPane);


	// установите наглядный диспетчер и стиль на основе постоянного значения
	OnApplicationLook(theApp.m_nAppLook);

	// Включить диалоговое окно расширенного управления окнами
	EnableWindowsDialog(ID_WINDOW_MANAGER, ID_WINDOW_MANAGER, TRUE);


	//DWMNCRENDERINGPOLICY policy = DWMNCRP_DISABLED;
	//DwmSetWindowAttribute(m_hWnd, DWMWA_ALLOW_NCPAINT, &policy, sizeof(policy));

	// Переключите порядок имени документа и имени приложения в заголовке окна. Это
	// повышает удобство использования панели задач, так как на эскизе отображается имя документа.
	ModifyStyle(0, FWS_PREFIXTITLE);

	return 0;
}

BOOL CMainFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	//cs.style = WS_POPUP | WS_THICKFRAME | WS_SYSMENU | WS_MAXIMIZEBOX | WS_MINIMIZEBOX;

	if( !CMDIFrameWndEx::PreCreateWindow(cs) )
		return FALSE;
	// TODO: изменить класс Window или стили посредством изменения
	//  CREATESTRUCT cs

	return TRUE;
}

BOOL CMainFrame::CreateDockingWindows()
{
	BOOL bNameValid;

	CString strTascPane = _T("Таск пейн");
	if (!m_TaskPane.Create(strTascPane, this, CRect(0, 0, 200, 200),
		TRUE, ID_TASKPANE, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_LEFT | CBRS_FLOAT_MULTI))
	{
		TRACE0("Не удалось создать окно \"Nfrcn\"\n");
	}

	// Создать представление классов
	CString strClassView;
	bNameValid = strClassView.LoadString(IDS_CLASS_VIEW);
	ASSERT(bNameValid);
	if (!m_wndClassView.Create(strClassView, this, CRect(0, 0, 200, 200), TRUE, ID_VIEW_CLASSVIEW, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_LEFT | CBRS_FLOAT_MULTI))
	{
		TRACE0("Не удалось создать окно \"Представление классов\"\n");
		return FALSE; // не удалось создать
	}

	// Создать представление файлов
	CString strFileView;
	bNameValid = strFileView.LoadString(IDS_FILE_VIEW);
	ASSERT(bNameValid);
	if (!m_wndFileView.Create(strFileView, this, CRect(0, 0, 200, 200), TRUE, ID_VIEW_FILEVIEW, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_LEFT| CBRS_FLOAT_MULTI))
	{
		TRACE0("Не удалось создать окно \"Представление файлов\"\n");
		return FALSE; // не удалось создать
	}

	// Создать окно вывода
	CString strOutputWnd;
	bNameValid = strOutputWnd.LoadString(IDS_OUTPUT_WND);
	ASSERT(bNameValid);
	if (!m_wndOutput.Create(strOutputWnd, this, CRect(0, 0, 100, 100), TRUE, ID_VIEW_OUTPUTWND, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_BOTTOM | CBRS_FLOAT_MULTI))
	{
		TRACE0("Не удалось создать окно \"Вывод\"\n");
		return FALSE; // не удалось создать
	}

	// Создать окно свойств
	CString strPropertiesWnd;
	bNameValid = strPropertiesWnd.LoadString(IDS_PROPERTIES_WND);
	ASSERT(bNameValid);
	if (!m_wndProperties.Create(strPropertiesWnd, this, CRect(0, 0, 200, 200), TRUE, ID_VIEW_PROPERTIESWND, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_RIGHT | CBRS_FLOAT_MULTI))
	{
		TRACE0("Не удалось создать окно \"Свойства\"\n");
		return FALSE; // не удалось создать
	}

	SetDockingWindowIcons(theApp.m_bHiColorIcons);
	return TRUE;
}

void CMainFrame::SetDockingWindowIcons(BOOL bHiColorIcons)
{
	HICON hFileViewIcon = (HICON) ::LoadImage(::AfxGetResourceHandle(), MAKEINTRESOURCE(bHiColorIcons ? IDI_FILE_VIEW_HC : IDI_FILE_VIEW), IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), 0);
	m_wndFileView.SetIcon(hFileViewIcon, FALSE);

	HICON hClassViewIcon = (HICON) ::LoadImage(::AfxGetResourceHandle(), MAKEINTRESOURCE(bHiColorIcons ? IDI_CLASS_VIEW_HC : IDI_CLASS_VIEW), IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), 0);
	m_wndClassView.SetIcon(hClassViewIcon, FALSE);

	HICON hOutputBarIcon = (HICON) ::LoadImage(::AfxGetResourceHandle(), MAKEINTRESOURCE(bHiColorIcons ? IDI_OUTPUT_WND_HC : IDI_OUTPUT_WND), IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), 0);
	m_wndOutput.SetIcon(hOutputBarIcon, FALSE);

	HICON hPropertiesBarIcon = (HICON) ::LoadImage(::AfxGetResourceHandle(), MAKEINTRESOURCE(bHiColorIcons ? IDI_PROPERTIES_WND_HC : IDI_PROPERTIES_WND), IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), 0);
	m_wndProperties.SetIcon(hPropertiesBarIcon, FALSE);

	UpdateMDITabbedBarsIcons();
}

BOOL CMainFrame::CreateCaptionBar()
{
	if (!m_wndCaptionBar.Create(WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS, this, ID_VIEW_CAPTION_BAR, -1, TRUE))
	{
		TRACE0("Не удалось создать заголовок окна\n");
		return FALSE;
	}

	BOOL bNameValid;

	CString strTemp, strTemp2;
	bNameValid = strTemp.LoadString(IDS_CAPTION_BUTTON);
	ASSERT(bNameValid);
	m_wndCaptionBar.SetButton(strTemp, ID_TOOLS_OPTIONS, CMFCCaptionBar::ALIGN_LEFT, FALSE);
	bNameValid = strTemp.LoadString(IDS_CAPTION_BUTTON_TIP);
	ASSERT(bNameValid);
	m_wndCaptionBar.SetButtonToolTip(strTemp);

	bNameValid = strTemp.LoadString(IDS_CAPTION_TEXT);
	ASSERT(bNameValid);
	m_wndCaptionBar.SetText(strTemp, CMFCCaptionBar::ALIGN_LEFT);

	m_wndCaptionBar.SetBitmap(IDB_INFO, RGB(255, 255, 255), FALSE, CMFCCaptionBar::ALIGN_LEFT);
	bNameValid = strTemp.LoadString(IDS_CAPTION_IMAGE_TIP);
	ASSERT(bNameValid);
	bNameValid = strTemp2.LoadString(IDS_CAPTION_IMAGE_TEXT);
	ASSERT(bNameValid);
	m_wndCaptionBar.SetImageToolTip(strTemp, strTemp2);

	return TRUE;
}

// Диагностика CMainFrame

#ifdef _DEBUG
void CMainFrame::AssertValid() const
{
	CMDIFrameWndEx::AssertValid();
}

void CMainFrame::Dump(CDumpContext& dc) const
{
	CMDIFrameWndEx::Dump(dc);
}
#endif //_DEBUG


// Обработчики сообщений CMainFrame

void CMainFrame::OnWindowManager()
{
	ShowWindowsDialog();
}

void CMainFrame::OnApplicationLook(UINT id)
{
	CWaitCursor wait;

	theApp.m_nAppLook = id;

		switch (theApp.m_nAppLook)
		{
		case ID_VIEW_APPLOOK_WINDOWS_7:
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows7));

			break;
		case ID_BUTTON_NOPSTYLE_DARK:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_DarkGray);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(FALSE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));


			break;
		case ID_BUTTON_NOPSTYLE_WHITE:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_White);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));


			break;
		case ID_BUTTON_NOPSTYLE_BLUE:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_Blue);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));


			break;
		case ID_BUTTON_NOPSTYLE_GREEN:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_Green);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));


			break;
		case ID_BUTTON_NOPSTYLE_PURPLE:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_Purple);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));


			break;
		case ID_BUTTON_NOPSTYLE_DARKBLUE:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_DarkBlue);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));

			break;
		case ID_BUTTON_NOPSTYLE_DARKGREEN:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_DarkGreen);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));

			break;
		case ID_BUTTON_NOPSTYLE_DARKYLLW:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_DarkYollow);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));

			break;
		case ID_BUTTON_NOPSTYLE_DARKPRPL:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_DarkPurple);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));

			break;
		case ID_BUTTON_NOPSTYLE_WHITEBLUE:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_WhiteBlue);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));

			break;
		case ID_BUTTON_NOPSTYLE_WHITEGREEN:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_WhiteGreen);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));

			break;
		case ID_BUTTON_NOPSTYLE_WHITEYLLW:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_WhiteYollow);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));

			break;
		case ID_BUTTON_NOPSTYLE_WHITEPRPL:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_WhitePurple);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));

			break;
		case ID_BUTTON_NOP_SYSCLR_LGHT:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_SystemColorLight);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));

			break;
		case ID_BUTTON_NOP_SYSCLR_DARK:
			CNPUP_VisualManager::SetStyle(CNPUP_VisualManager::nopStyle_SystemColorDark);
			CDockingManager::SetDockingMode(DT_SMART);
			m_wndRibbonBar.SetWindows7Look(TRUE);
			CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));

			break;

		}

		//CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CNPUP_VisualManager));
		CDockingManager::SetDockingMode(DT_SMART, AFX_SDT_VS2005);
		//m_wndRibbonBar.SetWindows7Look(FALSE);
	//}

	m_wndOutput.UpdateFonts();
	RedrawWindow(nullptr, nullptr, RDW_ALLCHILDREN | RDW_INVALIDATE | RDW_UPDATENOW | RDW_FRAME | RDW_ERASE);

	theApp.WriteInt(_T("ApplicationLook"), theApp.m_nAppLook);
}

void CMainFrame::OnUpdateApplicationLook(CCmdUI* pCmdUI)
{
	pCmdUI->SetRadio(theApp.m_nAppLook == pCmdUI->m_nID);
}

void CMainFrame::OnViewCaptionBar()
{
	m_wndCaptionBar.ShowWindow(m_wndCaptionBar.IsVisible() ? SW_HIDE : SW_SHOW);
	RecalcLayout(FALSE);
}

void CMainFrame::OnUpdateViewCaptionBar(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(m_wndCaptionBar.IsVisible());
}

void CMainFrame::OnOptions()
{
	CMFCRibbonCustomizeDialog *pOptionsDlg = new CMFCRibbonCustomizeDialog(this, &m_wndRibbonBar);
	ASSERT(pOptionsDlg != nullptr);

	pOptionsDlg->DoModal();
	delete pOptionsDlg;
}


void CMainFrame::OnSettingChange(UINT uFlags, LPCTSTR lpszSection)
{
	CMDIFrameWndEx::OnSettingChange(uFlags, lpszSection);
	m_wndOutput.UpdateFonts();
}
