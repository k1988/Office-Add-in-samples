// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

//#import "E:\\workspace_barmaco\\cpp\\hellocpp_iflytek\\NativeAddIn\\3rd_x86\\MSO.DLL" no_namespace
// CRibbonUI ��װ��

class CRibbonUI : public COleDispatchDriver
{
public:
	CRibbonUI(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CRibbonUI(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CRibbonUI(const CRibbonUI& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// IRibbonUI ����
public:
	void Invalidate()
	{
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void InvalidateControl(LPCTSTR ControlID)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ControlID);
	}
	void InvalidateControlMso(LPCTSTR ControlID)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ControlID);
	}
	void ActivateTab(LPCTSTR ControlID)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ControlID);
	}
	void ActivateTabMso(LPCTSTR ControlID)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ControlID);
	}
	void ActivateTabQ(LPCTSTR ControlID, LPCTSTR Namespace)
	{
		static BYTE parms[] = VTS_BSTR VTS_BSTR ;
		InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ControlID, Namespace);
	}

	// IRibbonUI ����
public:

};
