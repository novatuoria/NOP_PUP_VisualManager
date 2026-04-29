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

// NOP_PUP_VSMNGView.cpp: реализация класса CNOPPUPVSMNGView
//

#include "pch.h"
#include "framework.h"
// SHARED_HANDLERS можно определить в обработчиках фильтров просмотра реализации проекта ATL, эскизов
// и поиска; позволяет совместно использовать код документа в данным проекте.
#ifndef SHARED_HANDLERS
#include "NOP_PUP_VSMNG.h"
#endif

#include "NOP_PUP_VSMNGDoc.h"
#include "NOP_PUP_VSMNGView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CNOPPUPVSMNGView

IMPLEMENT_DYNCREATE(CNOPPUPVSMNGView, CScrollView)

BEGIN_MESSAGE_MAP(CNOPPUPVSMNGView, CScrollView)
	// Стандартные команды печати
	ON_COMMAND(ID_FILE_PRINT, &CScrollView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &CScrollView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CNOPPUPVSMNGView::OnFilePrintPreview)
	ON_WM_CONTEXTMENU()
	ON_WM_RBUTTONUP()
END_MESSAGE_MAP()

// Создание или уничтожение CNOPPUPVSMNGView

CNOPPUPVSMNGView::CNOPPUPVSMNGView() noexcept
{
	// TODO: добавьте код создания

}

CNOPPUPVSMNGView::~CNOPPUPVSMNGView()
{
}

BOOL CNOPPUPVSMNGView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: изменить класс Window или стили посредством изменения
	//  CREATESTRUCT cs

	return CScrollView::PreCreateWindow(cs);
}

// Рисование CNOPPUPVSMNGView

void CNOPPUPVSMNGView::OnDraw(CDC* pDC)
{
	CNOPPUPVSMNGDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	if (!pDoc)
		return;
	
	CRect rectWnd;
	//GetWindowRect(&rectWnd);

	pDC->GetClipBox(&rectWnd);
	pDC->FillSolidRect(rectWnd, RGB(128, 128, 128));

	// TODO: добавьте здесь код отрисовки для собственных данных
}

void CNOPPUPVSMNGView::OnInitialUpdate()
{
	CScrollView::OnInitialUpdate();

	CSize sizeTotal;
	// TODO: рассчитайте полный размер этого представления
	sizeTotal.cx = sizeTotal.cy = 100;
	SetScrollSizes(MM_TEXT, sizeTotal);
}


// Печать CNOPPUPVSMNGView


void CNOPPUPVSMNGView::OnFilePrintPreview()
{
#ifndef SHARED_HANDLERS
	AFXPrintPreview(this);
#endif
}

BOOL CNOPPUPVSMNGView::OnPreparePrinting(CPrintInfo* pInfo)
{
	// подготовка по умолчанию
	return DoPreparePrinting(pInfo);
}

void CNOPPUPVSMNGView::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: добавьте дополнительную инициализацию перед печатью
}

void CNOPPUPVSMNGView::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: добавьте очистку после печати
}

void CNOPPUPVSMNGView::OnRButtonUp(UINT /* nFlags */, CPoint point)
{
	ClientToScreen(&point);
	OnContextMenu(this, point);
}

void CNOPPUPVSMNGView::OnContextMenu(CWnd* /* pWnd */, CPoint point)
{
#ifndef SHARED_HANDLERS
	theApp.GetContextMenuManager()->ShowPopupMenu(IDR_POPUP_EDIT, point.x, point.y, this, TRUE);
#endif
}


// Диагностика CNOPPUPVSMNGView

#ifdef _DEBUG
void CNOPPUPVSMNGView::AssertValid() const
{
	CScrollView::AssertValid();
}

void CNOPPUPVSMNGView::Dump(CDumpContext& dc) const
{
	CScrollView::Dump(dc);
}

CNOPPUPVSMNGDoc* CNOPPUPVSMNGView::GetDocument() const // встроена неотлаженная версия
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CNOPPUPVSMNGDoc)));
	return (CNOPPUPVSMNGDoc*)m_pDocument;
}
#endif //_DEBUG


// Обработчики сообщений CNOPPUPVSMNGView
