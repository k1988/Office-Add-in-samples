// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

//#import "E:\\workspace_barmaco\\cpp\\hellocpp_iflytek\\NativeAddIn\\3rd_x86\\MSO.DLL" no_namespace
// CRibbonUI 包装类

class CRibbonUI : public COleDispatchDriver
{
public:
	CRibbonUI(){} // 调用 COleDispatchDriver 默认构造函数
	CRibbonUI(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CRibbonUI(const CRibbonUI& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// IRibbonUI 方法
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

	// IRibbonUI 属性
public:

};
