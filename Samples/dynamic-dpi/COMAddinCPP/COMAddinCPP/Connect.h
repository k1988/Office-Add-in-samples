// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Connect.h : Declaration of the CConnect

#pragma once
#include "resource.h"       // main symbols
#include "AddIn.h"
#include "CRibbonUI.h"

typedef IDispatchImpl<IRibbonExtensibility, &__uuidof(IRibbonExtensibility), &__uuidof(__Office), /* wMajor = */ 2, /* wMinor = */ 5>
RibbonImpl;

typedef IDispatchImpl<IRibbonCallback, &__uuidof(IRibbonCallback)> CallbackImpl;

// CConnect 实现_IDTExtensibility2接口，在插件加载退出或结束时调用，管理插件的生命周期
class ATL_NO_VTABLE CConnect :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConnect, &CLSID_Connect>,
	public IDispatchImpl<AddInDesignerObjects::_IDTExtensibility2, &AddInDesignerObjects::IID__IDTExtensibility2, &AddInDesignerObjects::LIBID_AddInDesignerObjects, 1, 0>,
	public IDispatchImpl<ICustomTaskPaneConsumer, &__uuidof(ICustomTaskPaneConsumer), &LIBID_Office, /* wMajor = */ 2, /* wMinor = */ 8>,
	public RibbonImpl,
	public CallbackImpl
{
public:
	CConnect()
	{
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_ADDIN)
	DECLARE_NOT_AGGREGATABLE(CConnect)

	BEGIN_COM_MAP(CConnect)
		COM_INTERFACE_ENTRY2(IDispatch, ICustomTaskPaneConsumer)
		COM_INTERFACE_ENTRY2(IDispatch, IRibbonCallback)
		COM_INTERFACE_ENTRY(AddInDesignerObjects::IDTExtensibility2)
		COM_INTERFACE_ENTRY(IRibbonExtensibility)
		COM_INTERFACE_ENTRY(IRibbonCallback)
		COM_INTERFACE_ENTRY(ICustomTaskPaneConsumer)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}


	HRESULT HrGetResource(int nId,
		LPCTSTR lpType,
		LPVOID* ppvResourceData,
		DWORD* pdwSizeInBytes)
	{
		HMODULE hModule = _AtlBaseModule.GetModuleInstance();
		if (!hModule)
			return E_UNEXPECTED;
		HRSRC hRsrc = FindResource(hModule, MAKEINTRESOURCE(nId), lpType);
		if (!hRsrc)
			return HRESULT_FROM_WIN32(GetLastError());
		HGLOBAL hGlobal = LoadResource(hModule, hRsrc);
		if (!hGlobal)
			return HRESULT_FROM_WIN32(GetLastError());
		*pdwSizeInBytes = SizeofResource(hModule, hRsrc);
		*ppvResourceData = LockResource(hGlobal);
		return S_OK;
	}

	BSTR GetXMLResource(int nId)
	{
		LPVOID pResourceData = NULL;
		DWORD dwSizeInBytes = 0;
		HRESULT hr = HrGetResource(nId, TEXT("XML"),
			&pResourceData, &dwSizeInBytes);
		if (FAILED(hr))
			return NULL;
		// Assumes that the data is not stored in Unicode.
		CComBSTR cbstr(dwSizeInBytes, reinterpret_cast<LPCSTR>(pResourceData));
		return cbstr.Detach();
	}

	// IRibbonExtensibility Methods
public:
	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR* RibbonXml)
	{
		if (!RibbonXml)
			return E_POINTER;
		*RibbonXml = GetXMLResource(IDR_XML1);
		return S_OK;
	}

	STDMETHOD(ButtonClicked)(IDispatch* ribbon)
	{
		MessageBoxW(NULL, L"Button Clicked!", L"NativeAddin", MB_OK);
		return S_OK;
	}

	STDMETHOD(UILoad)(IDispatch* ribbon) {
		m_ribbonUI = ribbon;

		//m_ribbonUI.InvalidateControl(_T("LoginButton"));
		return S_OK;
	}

	STDMETHOD(GetImage)(BSTR* pbstrImageId, IPictureDisp** ppdispImage);
	STDMETHOD(GetLabel)(IDispatch* ribbonCtrl, BSTR* RibbonXml);

public:
	STDMETHOD(LoginButtonClicked)(IDispatch* ribbon);
public:
	//IDTExtensibility2 implementation:
	STDMETHOD(OnConnection)(IDispatch * Application, AddInDesignerObjects::ext_ConnectMode ConnectMode, IDispatch *AddInInst, SAFEARRAY **custom);
	STDMETHOD(OnDisconnection)(AddInDesignerObjects::ext_DisconnectMode RemoveMode, SAFEARRAY **custom);
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom);
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom);
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom);

	CComPtr<IDispatch> m_pApplication;
	CComPtr<IDispatch> m_pAddInInstance;

	// ICustomTaskPaneConsumer Methods
public:
	STDMETHOD(CTPFactoryAvailable)(ICTPFactory * CTPFactoryInst);

public:
	CRibbonUI m_ribbonUI;
	static CComBSTR loginLable;
};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
