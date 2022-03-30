// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
// Connect.cpp : Implementation of CConnect
#include "stdafx.h"
#include "AddIn.h"
#include "Connect.h"
#include <string>

extern CAddInModule _AtlModule;

// When run, the Add-in wizard prepared the registry for the Add-in.
// At a later time, if the Add-in becomes unavailable for reasons such as:
//   1) You moved this project to a computer other than which is was originally created on.
//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
//   3) Registry corruption.
// you will need to re-register the Add-in by building the COMAddinCPPSetup project, 
// right click the project in the Solution Explorer, then choose install.


// CConnect
STDMETHODIMP CConnect::OnConnection(IDispatch *pApplication, AddInDesignerObjects::ext_ConnectMode /*ConnectMode*/, IDispatch *pAddInInst, SAFEARRAY ** /*custom*/ )
{
	pApplication->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pApplication);
	pAddInInst->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pAddInInstance);
	return S_OK;
}

STDMETHODIMP CConnect::OnDisconnection(AddInDesignerObjects::ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/ )
{
	m_pApplication = NULL;
	m_pAddInInstance = NULL;
	return S_OK;
}

STDMETHODIMP CConnect::OnAddInsUpdate (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CConnect::OnStartupComplete (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CConnect::OnBeginShutdown (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CConnect::CTPFactoryAvailable(ICTPFactory * CTPFactoryInst)
{

	_CustomTaskPane* pTaskPane = NULL;
	HRESULT hr = S_OK;

	VARIANTARG vargParentWindow;
	vargParentWindow.vt = VT_ERROR;
	vargParentWindow.scode = DISP_E_PARAMNOTFOUND;

	BSTR axControlID = L"COMAddinCPP.ATLControl";
	hr = CTPFactoryInst->CreateCTP(
		axControlID,
		CComBSTR(L"COM Add-in C++"), vargParentWindow, &pTaskPane);

	if (SUCCEEDED(hr))
	{
		hr = pTaskPane->put_Visible(TRUE);
	}

	return hr;
}


//CRibbonUI CConnect::m_ribbonUI = NULL;
CComBSTR CConnect::loginLable = CComBSTR("NotLogged");

STDMETHODIMP CConnect::GetImage(BSTR* pbstrImageId, IPictureDisp** ppdispImage)
{
	using namespace Gdiplus;
	int pngId(0);
	try
	{
		//test
		pngId = 204;
	}
	catch (...)
	{
		return E_UNEXPECTED;
	}
	LPVOID pResourceData = NULL;
	DWORD len = 0;
	HRESULT hr = HrGetResource(pngId, _T("PNG"), &pResourceData, &len);

	BYTE* lpRsrc = reinterpret_cast<BYTE*>(pResourceData);
	if (!lpRsrc)
	{
		return E_UNEXPECTED;
	}


	HGLOBAL m_hMem = GlobalAlloc(GMEM_FIXED, len);
	BYTE* pmem = (BYTE*)GlobalLock(m_hMem);
	memcpy(pmem, lpRsrc, len);
	GlobalUnlock(m_hMem);
	IStream* pstm;
	CreateStreamOnHGlobal(m_hMem, FALSE, &pstm);


	PICTDESC pic;
	memset(&pic, 0, sizeof pic);
	Bitmap* png = Bitmap::FromStream(pstm);
	HBITMAP hMap = NULL;
	png->GetHBITMAP(Color(), &hMap);
	pic.picType = PICTYPE_BITMAP;
	pic.bmp.hbitmap = hMap;


	OleCreatePictureIndirect(&pic, IID_IPictureDisp, true, (LPVOID*)ppdispImage);
	return S_OK;
}

STDMETHODIMP CConnect::GetLabel(IDispatch* ribbonCtrl, BSTR* RibbonXml)
{
	CComBSTR xml = loginLable.Copy();
	CComBSTR crlf = _T("\r\n");//add by ljm Ç¿ÖÆ»»ÐÐ
	xml.AppendBSTR(crlf);
	*RibbonXml = xml.Copy();
	return (*RibbonXml ? S_OK : E_OUTOFMEMORY);
}

STDMETHODIMP CConnect::LoginButtonClicked(IDispatch* ribbon)
{
	CConnect::loginLable = CComBSTR(_T("Logged"));
	m_ribbonUI.InvalidateControl(_T("LoginButton"));

	return S_OK;
}
