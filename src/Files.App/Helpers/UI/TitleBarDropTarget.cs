// Copyright (c) Files Community
// Licensed under the MIT License.

using Microsoft.Extensions.Logging;
using System.Runtime.InteropServices;
using Windows.Win32.Foundation;

namespace Files.App.Helpers
{
	/// <summary>
	/// Registers a drop target on the non client title bar window so that tabs and files can be
	/// dropped on the empty part of the title bar. WinUI 3 does not register an OLE drop target for
	/// its non client areas, which is why the drop has to be handled here.
	/// </summary>
	internal sealed class TitleBarDropTarget
	{
		private const ushort CF_HDROP = 15;
		private const uint DVASPECT_CONTENT = 1;
		private const int TYMED_HGLOBAL = 1;
		private const uint DROPEFFECT_NONE = 0;
		private const uint DROPEFFECT_COPY = 1;
		private const uint DROPEFFECT_MOVE = 2;
		private const int S_OK = 0;
		private const uint DragQueryFileCount = 0xFFFFFFFF;

		// DROPIMAGETYPE values used by the desktop drop description
		private const int DROPIMAGE_COPY = 2;
		private const int DROPIMAGE_MOVE = 3;

		private const uint CLSCTX_INPROC_SERVER = 1;
		private const uint GMEM_MOVEABLE = 0x0002;
		private const uint GMEM_ZEROINIT = 0x0040;

		private static readonly Guid ClsidDragDropHelper = new("4657278A-411B-11D2-839A-00C04FD918D0");
		private static readonly Guid IidDropTargetHelper = new("4657278B-411B-11D2-839A-00C04FD918D0");

		private readonly TabBar _tabBar;
		private GCHandle _selfHandle;
		private nint _dropTarget;
		private nint _registeredWindow;
		private nint _dropHelper;
		private nint _dataObject;
		private uint _currentEffect = DROPEFFECT_NONE;

		internal TitleBarDropTarget(TabBar tabBar)
		{
			_tabBar = tabBar;
		}

		/// <summary>
		/// Registers the drop target on the given title bar window. Calling it again for the same
		/// window is a no-op, which lets the caller invoke it whenever the drag region is refreshed.
		/// </summary>
		internal void TryRegister(HWND titleBarWindow)
		{
			nint window = titleBarWindow;

			if (window == 0 || window == _registeredWindow && _dropTarget != 0)
				return;

			try
			{
				if (_registeredWindow != 0 && _dropTarget != 0)
					Ole32.RevokeDragDrop(_registeredWindow);

				_dropTarget = DropTargetVtbl.Create(this, out _selfHandle);
				var result = Ole32.RegisterDragDrop(window, _dropTarget);

				if (result < 0)
				{
					App.Logger.LogWarning($"Failed to register the title bar drop target: 0x{result:X8}");
					return;
				}

				_registeredWindow = window;

				EnsureDropHelper();
			}
			catch (Exception ex)
			{
				App.Logger.LogWarning(ex, "Failed to register the title bar drop target");
			}
		}

		private uint GetEffect(nint dataObject)
		{
			// A tab of this process is being dragged, accept it as a move.
			if (TabBar.IsTabDragInProgress && TabBar.HasDraggedTabPayload)
				return DROPEFFECT_MOVE;

			return dataObject != 0 && HasFileDrop(dataObject) ? DROPEFFECT_COPY : DROPEFFECT_NONE;
		}

		private bool HasFileDrop(nint dataObject)
		{
			var format = CreateFormat();

			return DataObjectVtbl.GetDataObjectVtbl(dataObject).QueryGetData(dataObject, ref format) >= 0;
		}

		private static FORMATETC CreateFormat() => new()
		{
			cfFormat = CF_HDROP,
			dwAspect = DVASPECT_CONTENT,
			lindex = -1,
			tymed = TYMED_HGLOBAL
		};

		private void Drop(nint dataObject)
		{
			if (TabBar.TryGetDraggedTabPayload(out var serializedTab))
			{
				DropTab(serializedTab);
				return;
			}

			DropFiles(dataObject);
		}

		private void DropTab(string serializedTab)
		{
			// Tells TabView_TabDroppedOutside that the drop was already handled, so it must not
			// open a new window for the same tab.
			TabBar.MarkDropHandledOnTitleBar();

			// Started right away so the new tab is added before the source tab gets removed.
			_ = _tabBar.OpenTabFromTitleBarDropAsync(serializedTab);
		}

		private void DropFiles(nint dataObject)
		{
			var format = CreateFormat();
			var dataObjectVtbl = DataObjectVtbl.GetDataObjectVtbl(dataObject);

			if (dataObjectVtbl.GetData(dataObject, ref format, out var medium) < 0)
				return;

			try
			{
				var paths = ReadPaths(medium.unionmember);
				if (paths.Count is 0)
					return;

				// Dropping several files from the same folder must not open that folder more than once.
				var folders = paths
					.Select(path => SystemIO.Directory.Exists(path) ? path : SystemIO.Path.GetDirectoryName(path))
					.Where(path => !string.IsNullOrEmpty(path))
					.Distinct(StringComparer.OrdinalIgnoreCase)
					.ToArray();

				MainWindow.Instance.DispatcherQueue.TryEnqueue(async () =>
				{
					foreach (var folder in folders)
						await NavigationHelpers.OpenPathInNewTab(folder!, true);
				});
			}
			finally
			{
				Ole32.ReleaseStgMedium(ref medium);
			}
		}

		private static List<string> ReadPaths(nint hDrop)
		{
			var paths = new List<string>();

			if (hDrop == 0)
				return paths;

			var count = Shell32.DragQueryFile(hDrop, DragQueryFileCount, null, 0);
			for (uint index = 0; index < count; index++)
			{
				var length = Shell32.DragQueryFile(hDrop, index, null, 0);
				if (length is 0)
					continue;

				var buffer = new char[length + 1];
				if (Shell32.DragQueryFile(hDrop, index, buffer, (uint)buffer.Length) is 0)
					continue;

				var path = new string(buffer, 0, (int)length);
				if (!string.IsNullOrEmpty(path))
					paths.Add(path);
			}

			return paths;
		}

		/// <summary>
		/// Sets the drop description of the drag source, which is what draws the caption and the
		/// copy/move glyph next to the cursor.
		/// </summary>
		private void SetDropDescription(uint effect)
		{
			if (_dataObject == 0 || effect is DROPEFFECT_NONE)
				return;

			var formatName = "Desktop Drop Description";
			var cfFormat = User32.RegisterClipboardFormat(formatName);
			if (cfFormat is 0)
				return;

			var description = new DROPDESCRIPTION
			{
				type = effect is DROPEFFECT_MOVE ? DROPIMAGE_MOVE : DROPIMAGE_COPY,
				szMessage = effect is DROPEFFECT_MOVE
					? Strings.TabStripDragAndDropUIOverrideCaption.GetLocalizedResource()
					: Strings.OpenInNewTab.GetLocalizedResource(),
				szInsert = string.Empty
			};

			var format = new FORMATETC
			{
				cfFormat = (ushort)cfFormat,
				dwAspect = DVASPECT_CONTENT,
				lindex = -1,
				tymed = TYMED_HGLOBAL
			};

			var size = Marshal.SizeOf<DROPDESCRIPTION>();
			var handle = Kernel32.GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, (nuint)size);
			if (handle is 0)
				return;

			var pointer = Kernel32.GlobalLock(handle);
			if (pointer is 0)
			{
				Kernel32.GlobalFree(handle);
				return;
			}

			try
			{
				Marshal.StructureToPtr(description, pointer, false);
			}
			finally
			{
				Kernel32.GlobalUnlock(handle);
			}

			var medium = new STGMEDIUM
			{
				tymed = TYMED_HGLOBAL,
				unionmember = handle,
				pUnkForRelease = 0
			};

			// The data object takes ownership of the medium when the last argument is set.
			var setDataResult = DataObjectVtbl.GetDataObjectVtbl(_dataObject).SetData(_dataObject, ref format, ref medium, true);

			if (setDataResult < 0)
				Kernel32.GlobalFree(handle);
		}

		private void EnsureDropHelper()
		{
			if (_dropHelper != 0)
				return;

			try
			{
				var clsid = ClsidDragDropHelper;
				var iid = IidDropTargetHelper;

				if (Ole32.CoCreateInstance(ref clsid, 0, CLSCTX_INPROC_SERVER, ref iid, out var helper) >= 0)
					_dropHelper = helper;
			}
			catch (Exception ex)
			{
				App.Logger.LogWarning(ex, "Failed to create the drop target helper");
			}
		}

		private void HelpDragEnter(nint dataObject, POINTL point, uint effect)
		{
			EnsureDropHelper();

			if (_dropHelper == 0)
				return;

			DropHelperVtbl.Get(_dropHelper).DragEnter(_dropHelper, _registeredWindow, dataObject, ref point, effect);
		}

		private void HelpDragOver(POINTL point, uint effect)
		{
			if (_dropHelper == 0)
				return;

			DropHelperVtbl.Get(_dropHelper).DragOver(_dropHelper, ref point, effect);
		}

		private void HelpDragLeave()
		{
			if (_dropHelper == 0)
				return;

			DropHelperVtbl.Get(_dropHelper).DragLeave(_dropHelper);
		}

		private void HelpDrop(nint dataObject, POINTL point, uint effect)
		{
			if (_dropHelper == 0)
				return;

			DropHelperVtbl.Get(_dropHelper).Drop(_dropHelper, dataObject, ref point, effect);
		}

		private void SetDataObject(nint dataObject)
		{
			ReleaseDataObject();

			if (dataObject == 0)
				return;

			DataObjectVtbl.GetDataObjectVtbl(dataObject).AddRef(dataObject);
			_dataObject = dataObject;
		}

		private void ReleaseDataObject()
		{
			if (_dataObject == 0)
				return;

			DataObjectVtbl.GetDataObjectVtbl(_dataObject).Release(_dataObject);
			_dataObject = 0;
		}

		/// <summary>
		/// Builds a minimal IDropTarget COM object backed by a hand written vtable, so no COM
		/// interop or runtime marshaling support is required.
		/// </summary>
		private static class DropTargetVtbl
		{
			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate int QueryInterfaceDelegate(nint thisPtr, ref Guid riid, out nint interfacePtr);

			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate uint AddRefDelegate(nint thisPtr);

			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate uint ReleaseDelegate(nint thisPtr);

			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate int DragEnterDelegate(nint thisPtr, nint dataObject, uint keyState, POINTL point, ref uint effect);

			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate int DragOverDelegate(nint thisPtr, uint keyState, POINTL point, ref uint effect);

			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate int DragLeaveDelegate(nint thisPtr);

			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate int DropDelegate(nint thisPtr, nint dataObject, uint keyState, POINTL point, ref uint effect);

			// Keeps the delegates alive for the lifetime of the process.
			private static readonly QueryInterfaceDelegate _queryInterface = QueryInterface;
			private static readonly AddRefDelegate _addRef = AddRef;
			private static readonly ReleaseDelegate _release = Release;
			private static readonly DragEnterDelegate _dragEnter = DragEnter;
			private static readonly DragOverDelegate _dragOver = DragOver;
			private static readonly DragLeaveDelegate _dragLeave = DragLeave;
			private static readonly DropDelegate _drop = Drop;

			internal static nint Create(TitleBarDropTarget owner, out GCHandle handle)
			{
				var vtable = Marshal.AllocHGlobal(7 * IntPtr.Size);
				Marshal.WriteIntPtr(vtable, 0 * IntPtr.Size, Marshal.GetFunctionPointerForDelegate(_queryInterface));
				Marshal.WriteIntPtr(vtable, 1 * IntPtr.Size, Marshal.GetFunctionPointerForDelegate(_addRef));
				Marshal.WriteIntPtr(vtable, 2 * IntPtr.Size, Marshal.GetFunctionPointerForDelegate(_release));
				Marshal.WriteIntPtr(vtable, 3 * IntPtr.Size, Marshal.GetFunctionPointerForDelegate(_dragEnter));
				Marshal.WriteIntPtr(vtable, 4 * IntPtr.Size, Marshal.GetFunctionPointerForDelegate(_dragOver));
				Marshal.WriteIntPtr(vtable, 5 * IntPtr.Size, Marshal.GetFunctionPointerForDelegate(_dragLeave));
				Marshal.WriteIntPtr(vtable, 6 * IntPtr.Size, Marshal.GetFunctionPointerForDelegate(_drop));

				handle = GCHandle.Alloc(owner);

				var instance = Marshal.AllocHGlobal(2 * IntPtr.Size);
				Marshal.WriteIntPtr(instance, 0, vtable);
				Marshal.WriteIntPtr(instance, IntPtr.Size, GCHandle.ToIntPtr(handle));

				return instance;
			}

			private static TitleBarDropTarget GetOwner(nint thisPtr)
				=> (TitleBarDropTarget)GCHandle.FromIntPtr(Marshal.ReadIntPtr(thisPtr, IntPtr.Size)).Target!;

			private static int QueryInterface(nint thisPtr, ref Guid riid, out nint interfacePtr)
			{
				// IUnknown and IDropTarget are the only supported interfaces.
				interfacePtr = thisPtr;
				return S_OK;
			}

			private static uint AddRef(nint thisPtr) => 1;

			// The instance lives for the lifetime of the window, so it is never released.
			private static uint Release(nint thisPtr) => 1;

			private static int DragEnter(nint thisPtr, nint dataObject, uint keyState, POINTL point, ref uint effect)
			{
				var owner = GetOwner(thisPtr);

				try
				{
					owner._currentEffect = owner.GetEffect(dataObject);
					effect = owner._currentEffect;

					owner.SetDataObject(dataObject);
					owner.SetDropDescription(owner._currentEffect);
					owner.HelpDragEnter(dataObject, point, owner._currentEffect);
				}
				catch (Exception ex)
				{
					App.Logger.LogWarning(ex, "Failed to process a drag over the title bar");
				}

				return S_OK;
			}

			private static int DragOver(nint thisPtr, uint keyState, POINTL point, ref uint effect)
			{
				var owner = GetOwner(thisPtr);

				effect = owner._currentEffect;
				owner.HelpDragOver(point, effect);

				return S_OK;
			}

			private static int DragLeave(nint thisPtr)
			{
				var owner = GetOwner(thisPtr);

				owner._currentEffect = DROPEFFECT_NONE;
				owner.HelpDragLeave();
				owner.ReleaseDataObject();

				return S_OK;
			}

			private static int Drop(nint thisPtr, nint dataObject, uint keyState, POINTL point, ref uint effect)
			{
				var owner = GetOwner(thisPtr);

				effect = owner.GetEffect(dataObject);
				owner._currentEffect = DROPEFFECT_NONE;

				owner.HelpDrop(dataObject, point, effect);
				owner.ReleaseDataObject();

				if (effect is DROPEFFECT_NONE)
					return S_OK;

				owner.Drop(dataObject);

				return S_OK;
			}
		}

		/// <summary>
		/// The IDataObject members that are needed to check for, read and describe dropped data.
		/// </summary>
		private sealed class DataObjectVtbl
		{
			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate uint AddRefDelegate(nint thisPtr);

			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate uint ReleaseDelegate(nint thisPtr);

			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate int GetDataDelegate(nint thisPtr, ref FORMATETC format, out STGMEDIUM medium);

			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate int QueryGetDataDelegate(nint thisPtr, ref FORMATETC format);

			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate int SetDataDelegate(nint thisPtr, ref FORMATETC format, ref STGMEDIUM medium, bool release);

			private readonly AddRefDelegate _addRef;
			private readonly ReleaseDelegate _release;
			private readonly GetDataDelegate _getData;
			private readonly QueryGetDataDelegate _queryGetData;
			private readonly SetDataDelegate _setData;

			private DataObjectVtbl(nint dataObject)
			{
				var vtable = Marshal.ReadIntPtr(dataObject);
				_addRef = Marshal.GetDelegateForFunctionPointer<AddRefDelegate>(Marshal.ReadIntPtr(vtable, 1 * IntPtr.Size));
				_release = Marshal.GetDelegateForFunctionPointer<ReleaseDelegate>(Marshal.ReadIntPtr(vtable, 2 * IntPtr.Size));
				_getData = Marshal.GetDelegateForFunctionPointer<GetDataDelegate>(Marshal.ReadIntPtr(vtable, 3 * IntPtr.Size));
				_queryGetData = Marshal.GetDelegateForFunctionPointer<QueryGetDataDelegate>(Marshal.ReadIntPtr(vtable, 5 * IntPtr.Size));
				_setData = Marshal.GetDelegateForFunctionPointer<SetDataDelegate>(Marshal.ReadIntPtr(vtable, 7 * IntPtr.Size));
			}

			internal uint AddRef(nint dataObject) => _addRef(dataObject);

			internal uint Release(nint dataObject) => _release(dataObject);

			internal int GetData(nint dataObject, ref FORMATETC format, out STGMEDIUM medium)
				=> _getData(dataObject, ref format, out medium);

			internal int QueryGetData(nint dataObject, ref FORMATETC format)
				=> _queryGetData(dataObject, ref format);

			internal int SetData(nint dataObject, ref FORMATETC format, ref STGMEDIUM medium, bool release)
				=> _setData(dataObject, ref format, ref medium, release);

			internal static DataObjectVtbl GetDataObjectVtbl(nint dataObject) => new(dataObject);
		}

		/// <summary>
		/// The IDropTargetHelper members used to draw the standard drag image.
		/// </summary>
		private sealed class DropHelperVtbl
		{
			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate int DragEnterDelegate(nint thisPtr, nint window, nint dataObject, ref POINTL point, uint effect);

			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate int DragOverDelegate(nint thisPtr, ref POINTL point, uint effect);

			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate int DragLeaveDelegate(nint thisPtr);

			[UnmanagedFunctionPointer(CallingConvention.Winapi)]
			private delegate int DropDelegate(nint thisPtr, nint dataObject, ref POINTL point, uint effect);

			private readonly DragEnterDelegate _dragEnter;
			private readonly DragOverDelegate _dragOver;
			private readonly DragLeaveDelegate _dragLeave;
			private readonly DropDelegate _drop;

			private DropHelperVtbl(nint helper)
			{
				var vtable = Marshal.ReadIntPtr(helper);
				_dragEnter = Marshal.GetDelegateForFunctionPointer<DragEnterDelegate>(Marshal.ReadIntPtr(vtable, 3 * IntPtr.Size));
				_dragOver = Marshal.GetDelegateForFunctionPointer<DragOverDelegate>(Marshal.ReadIntPtr(vtable, 4 * IntPtr.Size));
				_dragLeave = Marshal.GetDelegateForFunctionPointer<DragLeaveDelegate>(Marshal.ReadIntPtr(vtable, 5 * IntPtr.Size));
				_drop = Marshal.GetDelegateForFunctionPointer<DropDelegate>(Marshal.ReadIntPtr(vtable, 6 * IntPtr.Size));
			}

			internal int DragEnter(nint helper, nint window, nint dataObject, ref POINTL point, uint effect)
				=> _dragEnter(helper, window, dataObject, ref point, effect);

			internal int DragOver(nint helper, ref POINTL point, uint effect)
				=> _dragOver(helper, ref point, effect);

			internal int DragLeave(nint helper) => _dragLeave(helper);

			internal int Drop(nint helper, nint dataObject, ref POINTL point, uint effect)
				=> _drop(helper, dataObject, ref point, effect);

			internal static DropHelperVtbl Get(nint helper) => new(helper);
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct POINTL
		{
			public int x;
			public int y;
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct FORMATETC
		{
			public ushort cfFormat;
			public nint ptd;
			public uint dwAspect;
			public int lindex;
			public int tymed;
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct STGMEDIUM
		{
			public int tymed;
			public nint unionmember;
			public nint pUnkForRelease;
		}

		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
		private struct DROPDESCRIPTION
		{
			public int type;

			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
			public string szMessage;

			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
			public string szInsert;
		}

		private static class Ole32
		{
			[DllImport("ole32.dll")]
			internal static extern int RegisterDragDrop(nint window, nint dropTarget);

			[DllImport("ole32.dll")]
			internal static extern int RevokeDragDrop(nint window);

			[DllImport("ole32.dll")]
			internal static extern void ReleaseStgMedium(ref STGMEDIUM medium);

			[DllImport("ole32.dll")]
			internal static extern int CoCreateInstance(ref Guid clsid, nint outer, uint context, ref Guid iid, out nint instance);
		}

		private static class Shell32
		{
			[DllImport("shell32.dll", CharSet = CharSet.Unicode, EntryPoint = "DragQueryFileW")]
			internal static extern uint DragQueryFile(nint hDrop, uint fileIndex, char[]? fileName, uint fileNameLength);
		}

		private static class User32
		{
			[DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "RegisterClipboardFormatW")]
			internal static extern uint RegisterClipboardFormat(string format);
		}

		private static class Kernel32
		{
			[DllImport("kernel32.dll")]
			internal static extern nint GlobalAlloc(uint flags, nuint bytes);

			[DllImport("kernel32.dll")]
			internal static extern nint GlobalLock(nint memory);

			[DllImport("kernel32.dll")]
			internal static extern bool GlobalUnlock(nint memory);

			[DllImport("kernel32.dll")]
			internal static extern nint GlobalFree(nint memory);
		}
	}
}
