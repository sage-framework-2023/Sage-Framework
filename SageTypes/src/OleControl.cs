using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;

namespace SageTypes
{
    // Base dos controles do Sage para UserForms (Caixa de ferramentas > Controles adicionais).
    //
    // O protocolo ActiveX é implementado aqui, à mão, como faz o ATL do C++: o suporte a ActiveX
    // que vem no Control do WinForms derruba o FM20.DLL (o motor dos UserForms) logo ao inserir o
    // controle. A classe COM não é um Control: ela cria uma janela (OleHost, com o conteúdo do
    // WinForms dentro) quando o UserForm a ativa, e desenha sozinha (IViewObject) quando não há
    // janela.
    //
    // Cada controle (TextBox...) implementa:
    //   CreateContent    o controle do WinForms que fica dentro da janela
    //   Paint            o desenho sem janela (designer, impressão) e o fundo e a borda da janela
    //   SaveProperties / LoadProperties / InitProperties   as propriedades gravadas no arquivo
    //
    // No designer (UserMode = False) a janela deixa os cliques passarem para o UserForm
    // (HTTRANSPARENT), para selecionar, mover e redimensionar o controle.
    [ComVisible(false)]
    public abstract class OleControl : IOleObject, IOleInPlaceObject, IOleInPlaceActiveObject, IOleControl,
        IViewObject, IViewObject2, IPersist, IPersistStreamInit, IPersistStream, IObjectSafety, ICustomQueryInterface
    {
        const int S_OK = 0, S_FALSE = 1;
        const int E_NOTIMPL = unchecked((int)0x80004001), E_FAIL = unchecked((int)0x80004005), E_INVALIDARG = unchecked((int)0x80070057);
        const int OLEOBJ_S_INVALIDVERB = 0x00040181, OLE_E_NOCONNECTION = unchecked((int)0x80040004);
        const int OLEIVERB_PRIMARY = 0, OLEIVERB_SHOW = -1, OLEIVERB_OPEN = -2, OLEIVERB_HIDE = -3, OLEIVERB_UIACTIVATE = -4,
            OLEIVERB_INPLACEACTIVATE = -5, OLEIVERB_DISCARDUNDOSTATE = -6, OLEIVERB_PROPERTIES = -7;
        const int DVASPECT_CONTENT = 1;
        const int DISPID_AMBIENT_USERMODE = -709, DISPID_UNKNOWN = -1;
        // RECOMPOSEONRESIZE | CANTLINKINSIDE | INSIDEOUT | ACTIVATEWHENVISIBLE | SETCLIENTSITEFIRST (= MiscStatus no registro)
        internal const int MiscStatus = 0x1 | 0x10 | 0x80 | 0x100 | 0x20000;
        const uint Magic = 0x53414745; // "SAGE"

        IOleClientSite site;
        IOleInPlaceSite inPlaceSite;
        IOleControlSite controlSite;
        IOleInPlaceFrame frame;
        IOleInPlaceUIWindow document;
        IAdviseSink viewSink;
        readonly Dictionary<int, IAdviseSink> sinks = new Dictionary<int, IAdviseSink>();
        int nextCookie = 1, freezeEvents;
        SIZEL extent;
        bool inPlace, uiActive, activating, dirty, userMode = true;
        OleHost host;
        Control content;

        // Recusa IMarshal e IAgileObject: o .NET responderia com o empacotador livre de threads, e
        // o UserForm, que embrulha o controle (extender), repassaria a pergunta a ele; fora do
        // processo (automação) chegaria o controle cru, sem o embrulho, e o FM20.DLL cai.
        // Recusando, o COM empacota o embrulho pelo caminho padrão.
        static readonly Guid IID_IMarshal = new Guid("00000003-0000-0000-C000-000000000046");
        static readonly Guid IID_IAgileObject = new Guid("94EA2B94-E9CC-49E0-C0FF-EE64CA8F5B90");

        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;
            return iid == IID_IMarshal || iid == IID_IAgileObject ? CustomQueryInterfaceResult.Failed : CustomQueryInterfaceResult.NotHandled;
        }

        protected OleControl(int widthPoints, int heightPoints)
        {
            // HIMETRIC: 1 pt = 2540 / 72
            extent.cx = widthPoints * 2540 / 72;
            extent.cy = heightPoints * 2540 / 72;
        }

        protected abstract Control CreateContent();
        // windowed: a janela do controle existe (o conteúdo do WinForms desenha o resto)
        protected abstract void Paint(Graphics g, Rectangle bounds, bool windowed);
        protected abstract void SaveProperties(IDictionary<string, object> properties);
        protected abstract void LoadProperties(IDictionary<string, object> properties);
        protected abstract void InitProperties();
        protected abstract string TypeName { get; }

        // False no designer do UserForm
        protected bool UserMode { get { return userMode; } }
        // Eventos para o VBA: só com o form rodando e fora de FreezeEvents
        protected bool EventsEnabled { get { return userMode && freezeEvents == 0; } }
        protected Control Content { get { return content; } }
        protected OleHost Host { get { return host; } }

        // Uma propriedade mudou: marca para gravar e redesenha
        protected void Changed()
        {
            dirty = true;
            if (host != null) host.Invalidate(true);
            if (viewSink != null)
            {
                try { viewSink.OnViewChange(DVASPECT_CONTENT, -1); } catch (Exception) { }
            }
        }

        // Teclas que o UserForm trata (Tab, Enter, Esc): devolve True se ele usou
        protected bool ForwardKey(ref Message m)
        {
            if (controlSite == null || !userMode) return false;
            OLEMSG msg = new OLEMSG { hwnd = m.HWnd, message = m.Msg, wParam = m.WParam, lParam = m.LParam };
            int modifiers = (Control.ModifierKeys & Keys.Shift) != 0 ? 1 : 0;
            if ((Control.ModifierKeys & Keys.Control) != 0) modifiers |= 2;
            if ((Control.ModifierKeys & Keys.Alt) != 0) modifiers |= 4;
            try { return controlSite.TranslateAccelerator(ref msg, modifiers) == S_OK; }
            catch (Exception) { return false; }
        }

        // O conteúdo ganhou ou perdeu o foco (clique, Tab): o UserForm precisa saber, para
        // trocar o controle ativo e disparar Enter/Exit
        protected void ContentFocus(bool got)
        {
            if (got && inPlace && !uiActive && !activating) UIActivate(false);
            if (controlSite != null)
            {
                try { controlSite.OnFocus(got ? 1 : 0); } catch (Exception) { }
            }
        }

        // ------------------------------------------------------------------
        // Janela
        // ------------------------------------------------------------------

        void EnsureHost(IntPtr parent)
        {
            if (host == null)
            {
                host = new OleHost(this);
                content = CreateContent();
                host.Controls.Add(content);
            }
            host.Transparent = !userMode;
            if (!host.IsHandleCreated)
            {
                host.ParentHandle = parent;
                host.CreateControl();
            }
            else if (GetParent(host.Handle) != parent) SetParent(host.Handle, parent);
        }

        void Position(RECT pos)
        {
            if (host == null || !host.IsHandleCreated) return;
            SetWindowPos(host.Handle, IntPtr.Zero, pos.Left, pos.Top, Math.Max(0, pos.Right - pos.Left), Math.Max(0, pos.Bottom - pos.Top),
                SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);
        }

        internal void PaintHost(Graphics g, Rectangle bounds) { Paint(g, bounds, true); }

        int InPlaceActivate(bool ui, IntPtr posRect)
        {
            if (site == null) return E_FAIL;
            if (inPlaceSite == null) inPlaceSite = site as IOleInPlaceSite;
            if (inPlaceSite == null) return E_FAIL;
            if (!inPlace)
            {
                if (inPlaceSite.CanInPlaceActivate() != S_OK) return E_FAIL;
                inPlaceSite.OnInPlaceActivate();
                inPlace = true;
                IntPtr parent;
                inPlaceSite.GetWindow(out parent);
                OLEINPLACEFRAMEINFO info = new OLEINPLACEFRAMEINFO();
                info.cb = Marshal.SizeOf(typeof(OLEINPLACEFRAMEINFO));
                RECT pos, clip;
                ReleaseFrame();
                inPlaceSite.GetWindowContext(out frame, out document, out pos, out clip, ref info);
                if (posRect != IntPtr.Zero) pos = (RECT)Marshal.PtrToStructure(posRect, typeof(RECT));
                EnsureHost(parent);
                Position(pos);
            }
            if (ui && !uiActive) UIActivate(true);
            return S_OK;
        }

        void UIActivate(bool focus)
        {
            uiActive = true;
            activating = true;
            try
            {
                inPlaceSite.OnUIActivate();
                if (frame != null) frame.SetActiveObject(this, null);
                if (document != null) document.SetActiveObject(this, null);
                if (focus && userMode && content != null && !content.ContainsFocus) content.Focus();
            }
            catch (Exception) { }
            finally { activating = false; }
        }

        void ReleaseFrame()
        {
            if (frame != null) { Marshal.ReleaseComObject(frame); frame = null; }
            if (document != null) { Marshal.ReleaseComObject(document); document = null; }
        }

        void ReleaseSites()
        {
            if (inPlaceSite != null && Marshal.IsComObject(inPlaceSite)) Marshal.ReleaseComObject(inPlaceSite);
            inPlaceSite = null;
            controlSite = null;
            if (site != null && Marshal.IsComObject(site)) Marshal.ReleaseComObject(site);
            site = null;
        }

        void ReadAmbient()
        {
            userMode = true;
            if (site == null) return;
            try
            {
                object value = site.GetType().InvokeMember("[DispID=" + DISPID_AMBIENT_USERMODE + "]", BindingFlags.GetProperty, null, site, null);
                if (value is bool) userMode = (bool)value;
            }
            catch (Exception) { }
            if (host != null) host.Transparent = !userMode;
        }

        // ------------------------------------------------------------------
        // IOleObject
        // ------------------------------------------------------------------

        int IOleObject.SetClientSite(IOleClientSite clientSite)
        {
            if (clientSite == null)
            {
                ReleaseSites();
                return S_OK;
            }
            site = clientSite;
            controlSite = clientSite as IOleControlSite;
            ReadAmbient();
            return S_OK;
        }

        int IOleObject.GetClientSite(out IOleClientSite clientSite) { clientSite = site; return S_OK; }
        int IOleObject.SetHostNames(string app, string obj) { return S_OK; }

        int IOleObject.Close(int saveOption)
        {
            ((IOleInPlaceObject)this).InPlaceDeactivate();
            foreach (IAdviseSink sink in sinks.Values)
            {
                try { sink.OnClose(); } catch (Exception) { }
            }
            if (host != null)
            {
                host.Dispose();
                host = null;
                content = null;
            }
            return S_OK;
        }

        int IOleObject.SetMoniker(int which, IntPtr moniker) { return E_NOTIMPL; }
        int IOleObject.GetMoniker(int assign, int which, out IntPtr moniker) { moniker = IntPtr.Zero; return E_NOTIMPL; }
        int IOleObject.InitFromData(IntPtr data, int creation, int reserved) { return E_NOTIMPL; }
        int IOleObject.GetClipboardData(int reserved, out IntPtr data) { data = IntPtr.Zero; return E_NOTIMPL; }

        int IOleObject.DoVerb(int verb, IntPtr msg, IOleClientSite activeSite, int index, IntPtr parent, IntPtr posRect)
        {
            switch (verb)
            {
                case OLEIVERB_SHOW:
                case OLEIVERB_INPLACEACTIVATE:
                    return InPlaceActivate(false, posRect);
                case OLEIVERB_PRIMARY:
                case OLEIVERB_UIACTIVATE:
                    return InPlaceActivate(true, posRect);
                case OLEIVERB_HIDE:
                    ((IOleInPlaceObject)this).UIDeactivate();
                    if (host != null && host.IsHandleCreated) ShowWindow(host.Handle, SW_HIDE);
                    return S_OK;
                case OLEIVERB_OPEN:
                case OLEIVERB_PROPERTIES:
                case OLEIVERB_DISCARDUNDOSTATE:
                    return E_NOTIMPL;
                default:
                    return verb > 0 ? OLEOBJ_S_INVALIDVERB : E_NOTIMPL;
            }
        }

        int IOleObject.EnumVerbs(out IntPtr enumerator) { enumerator = IntPtr.Zero; return E_NOTIMPL; }
        int IOleObject.Update() { return S_OK; }
        int IOleObject.IsUpToDate() { return S_OK; }
        int IOleObject.GetUserClassID(out Guid clsid) { clsid = GetType().GUID; return S_OK; }
        int IOleObject.GetUserType(int form, out string type) { type = TypeName; return S_OK; }

        int IOleObject.SetExtent(int aspect, ref SIZEL size)
        {
            if (aspect != DVASPECT_CONTENT) return E_FAIL;
            if (extent.cx != size.cx || extent.cy != size.cy) dirty = true;
            extent = size;
            return S_OK;
        }

        int IOleObject.GetExtent(int aspect, out SIZEL size)
        {
            size = extent;
            return aspect == DVASPECT_CONTENT ? S_OK : E_FAIL;
        }

        int IOleObject.Advise(IAdviseSink sink, out int cookie)
        {
            cookie = 0;
            if (sink == null) return E_INVALIDARG;
            cookie = nextCookie++;
            sinks[cookie] = sink;
            return S_OK;
        }

        int IOleObject.Unadvise(int cookie) { return sinks.Remove(cookie) ? S_OK : OLE_E_NOCONNECTION; }
        int IOleObject.EnumAdvise(out IntPtr enumerator) { enumerator = IntPtr.Zero; return E_NOTIMPL; }
        int IOleObject.GetMiscStatus(int aspect, out int status) { status = MiscStatus; return S_OK; }
        int IOleObject.SetColorScheme(IntPtr palette) { return E_NOTIMPL; }

        // ------------------------------------------------------------------
        // IOleInPlaceObject / IOleInPlaceActiveObject
        // ------------------------------------------------------------------

        int GetWindow(out IntPtr hwnd)
        {
            hwnd = host != null && host.IsHandleCreated ? host.Handle : IntPtr.Zero;
            return hwnd != IntPtr.Zero ? S_OK : E_FAIL;
        }

        int IOleInPlaceObject.GetWindow(out IntPtr hwnd) { return GetWindow(out hwnd); }
        int IOleInPlaceObject.ContextSensitiveHelp(int enter) { return E_NOTIMPL; }

        int IOleInPlaceObject.InPlaceDeactivate()
        {
            if (!inPlace) return S_OK;
            ((IOleInPlaceObject)this).UIDeactivate();
            inPlace = false;
            if (host != null && host.IsHandleCreated) ShowWindow(host.Handle, SW_HIDE);
            try { if (inPlaceSite != null) inPlaceSite.OnInPlaceDeactivate(); } catch (Exception) { }
            ReleaseFrame();
            return S_OK;
        }

        int IOleInPlaceObject.UIDeactivate()
        {
            if (!uiActive) return S_OK;
            uiActive = false;
            try
            {
                if (frame != null) frame.SetActiveObject(null, null);
                if (document != null) document.SetActiveObject(null, null);
                if (inPlaceSite != null) inPlaceSite.OnUIDeactivate(0);
            }
            catch (Exception) { }
            return S_OK;
        }

        int IOleInPlaceObject.SetObjectRects(ref RECT pos, ref RECT clip) { Position(pos); return S_OK; }
        int IOleInPlaceObject.ReactivateAndUndo() { return E_NOTIMPL; }

        int IOleInPlaceActiveObject.GetWindow(out IntPtr hwnd) { return GetWindow(out hwnd); }
        int IOleInPlaceActiveObject.ContextSensitiveHelp(int enter) { return E_NOTIMPL; }
        int IOleInPlaceActiveObject.TranslateAccelerator(ref OLEMSG msg) { return S_FALSE; }
        int IOleInPlaceActiveObject.OnFrameWindowActivate(int activate) { return S_OK; }
        int IOleInPlaceActiveObject.OnDocWindowActivate(int activate) { return S_OK; }
        int IOleInPlaceActiveObject.ResizeBorder(ref RECT border, IntPtr window, int frameWindow) { return S_OK; }
        int IOleInPlaceActiveObject.EnableModeless(int enable) { return S_OK; }

        // ------------------------------------------------------------------
        // IOleControl
        // ------------------------------------------------------------------

        int IOleControl.GetControlInfo(IntPtr info)
        {
            // CONTROLINFO sem aceleradores: zera tudo depois do cb
            int size = Marshal.ReadInt32(info);
            for (int i = 4; i < size; i++) Marshal.WriteByte(info, i, 0);
            return S_OK;
        }

        int IOleControl.OnMnemonic(ref OLEMSG msg) { return E_NOTIMPL; }

        int IOleControl.OnAmbientPropertyChange(int dispid)
        {
            if (dispid == DISPID_AMBIENT_USERMODE || dispid == DISPID_UNKNOWN) ReadAmbient();
            return S_OK;
        }

        int IOleControl.FreezeEvents(int freeze)
        {
            freezeEvents = Math.Max(0, freezeEvents + (freeze != 0 ? 1 : -1));
            return S_OK;
        }

        // ------------------------------------------------------------------
        // IViewObject / IViewObject2: desenho sem janela
        // ------------------------------------------------------------------

        int Draw(int aspect, IntPtr hdc, IntPtr bounds)
        {
            if (aspect != DVASPECT_CONTENT) return E_NOTIMPL;
            if (hdc == IntPtr.Zero || bounds == IntPtr.Zero) return E_INVALIDARG;
            RECT r = (RECT)Marshal.PtrToStructure(bounds, typeof(RECT));
            try
            {
                using (Graphics g = Graphics.FromHdc(hdc))
                    Paint(g, Rectangle.FromLTRB(r.Left, r.Top, r.Right, r.Bottom), false);
            }
            catch (Exception) { return E_FAIL; }
            return S_OK;
        }

        int IViewObject.Draw(int aspect, int index, IntPtr aspectInfo, IntPtr device, IntPtr targetDc, IntPtr hdc, IntPtr bounds,
            IntPtr wbounds, IntPtr cont, IntPtr contArg) { return Draw(aspect, hdc, bounds); }
        int IViewObject.GetColorSet(int aspect, int index, IntPtr aspectInfo, IntPtr device, IntPtr targetDc, out IntPtr palette)
        { palette = IntPtr.Zero; return E_NOTIMPL; }
        int IViewObject.Freeze(int aspect, int index, IntPtr aspectInfo, out int freeze) { freeze = 0; return E_NOTIMPL; }
        int IViewObject.Unfreeze(int freeze) { return E_NOTIMPL; }
        int IViewObject.SetAdvise(int aspects, int advf, IAdviseSink sink) { viewSink = sink; return S_OK; }
        int IViewObject.GetAdvise(out int aspects, out int advf, out IAdviseSink sink) { aspects = DVASPECT_CONTENT; advf = 0; sink = viewSink; return S_OK; }

        int IViewObject2.Draw(int aspect, int index, IntPtr aspectInfo, IntPtr device, IntPtr targetDc, IntPtr hdc, IntPtr bounds,
            IntPtr wbounds, IntPtr cont, IntPtr contArg) { return Draw(aspect, hdc, bounds); }
        int IViewObject2.GetColorSet(int aspect, int index, IntPtr aspectInfo, IntPtr device, IntPtr targetDc, out IntPtr palette)
        { palette = IntPtr.Zero; return E_NOTIMPL; }
        int IViewObject2.Freeze(int aspect, int index, IntPtr aspectInfo, out int freeze) { freeze = 0; return E_NOTIMPL; }
        int IViewObject2.Unfreeze(int freeze) { return E_NOTIMPL; }
        int IViewObject2.SetAdvise(int aspects, int advf, IAdviseSink sink) { viewSink = sink; return S_OK; }
        int IViewObject2.GetAdvise(out int aspects, out int advf, out IAdviseSink sink) { aspects = DVASPECT_CONTENT; advf = 0; sink = viewSink; return S_OK; }
        int IViewObject2.GetExtent(int aspect, int index, IntPtr device, out SIZEL size) { size = extent; return S_OK; }

        // ------------------------------------------------------------------
        // Gravação no arquivo (IPersistStreamInit): "SAGE", versão 1, tamanho, propriedades
        // (nome, tipo, valor). Propriedades desconhecidas são ignoradas ao ler, então
        // versões novas do controle podem acrescentar propriedades.
        // ------------------------------------------------------------------

        int GetClassID(out Guid clsid) { clsid = GetType().GUID; return S_OK; }
        int IPersist.GetClassID(out Guid clsid) { return GetClassID(out clsid); }
        int IPersistStreamInit.GetClassID(out Guid clsid) { return GetClassID(out clsid); }
        int IPersistStream.GetClassID(out Guid clsid) { return GetClassID(out clsid); }
        int IPersistStreamInit.IsDirty() { return dirty ? S_OK : S_FALSE; }
        int IPersistStream.IsDirty() { return dirty ? S_OK : S_FALSE; }
        int IPersistStreamInit.Load(IStream stream) { return Load(stream); }
        int IPersistStream.Load(IStream stream) { return Load(stream); }
        int IPersistStreamInit.Save(IStream stream, int clearDirty) { return Save(stream, clearDirty != 0); }
        int IPersistStream.Save(IStream stream, int clearDirty) { return Save(stream, clearDirty != 0); }
        int IPersistStreamInit.GetSizeMax(out long size) { size = Serialize().Length; return S_OK; }
        int IPersistStream.GetSizeMax(out long size) { size = Serialize().Length; return S_OK; }

        int IPersistStreamInit.InitNew()
        {
            InitProperties();
            dirty = false;
            return S_OK;
        }

        int Save(IStream stream, bool clearDirty)
        {
            try
            {
                byte[] bytes = Serialize();
                stream.Write(bytes, bytes.Length, IntPtr.Zero);
                if (clearDirty) dirty = false;
                return S_OK;
            }
            catch (Exception) { return E_FAIL; }
            finally { Release(stream); }
        }

        // Solta o stream na hora, como um controle nativo: o .NET o manteria aberto até o coletor
        // de lixo passar, e o form, ao rodar, abre o mesmo stream ("o") em modo exclusivo; com ele
        // ainda aberto, dá STG_E_ACCESSDENIED ("Erro de acesso a caminho/arquivo")
        static void Release(object com)
        {
            if (com != null && Marshal.IsComObject(com)) Marshal.FinalReleaseComObject(com);
        }

        byte[] Serialize()
        {
            {
                Dictionary<string, object> properties = new Dictionary<string, object>();
                SaveProperties(properties);
                // O tamanho vai junto, como nos controles nativos: o UserForm o pede (GetExtent)
                // depois de carregar
                properties["_Width"] = extent.cx;
                properties["_Height"] = extent.cy;
                MemoryStream body = new MemoryStream();
                using (BinaryWriter w = new BinaryWriter(body, System.Text.Encoding.UTF8))
                {
                    w.Write(properties.Count);
                    foreach (KeyValuePair<string, object> p in properties)
                    {
                        w.Write(p.Key);
                        if (p.Value is string) { w.Write((byte)0); w.Write((string)p.Value); }
                        else if (p.Value is bool) { w.Write((byte)1); w.Write((bool)p.Value); }
                        else if (p.Value is int) { w.Write((byte)2); w.Write((int)p.Value); }
                        else if (p.Value is double) { w.Write((byte)3); w.Write((double)p.Value); }
                        else throw new InvalidOperationException(p.Key);
                    }
                }
                byte[] data = body.ToArray();
                MemoryStream all = new MemoryStream();
                using (BinaryWriter w = new BinaryWriter(all))
                {
                    w.Write(Magic);
                    w.Write(1);
                    w.Write(data.Length);
                    w.Write(data);
                }
                return all.ToArray();
            }
        }

        int Load(IStream stream)
        {
            try
            {
                byte[] header = Read(stream, 12);
                if (BitConverter.ToUInt32(header, 0) != Magic) return E_FAIL;
                byte[] data = Read(stream, BitConverter.ToInt32(header, 8));
                Dictionary<string, object> properties = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                using (BinaryReader r = new BinaryReader(new MemoryStream(data), System.Text.Encoding.UTF8))
                {
                    int count = r.ReadInt32();
                    for (int i = 0; i < count; i++)
                    {
                        string name = r.ReadString();
                        byte type = r.ReadByte();
                        object value;
                        switch (type)
                        {
                            case 0: value = r.ReadString(); break;
                            case 1: value = r.ReadBoolean(); break;
                            case 2: value = r.ReadInt32(); break;
                            case 3: value = r.ReadDouble(); break;
                            default: throw new InvalidDataException(name);
                        }
                        properties[name] = value;
                    }
                }
                InitProperties();
                LoadProperties(properties);
                extent.cx = Get(properties, "_Width", extent.cx);
                extent.cy = Get(properties, "_Height", extent.cy);
                dirty = false;
                return S_OK;
            }
            catch (Exception) { return E_FAIL; }
            finally { Release(stream); }
        }

        static byte[] Read(IStream stream, int count)
        {
            byte[] buffer = new byte[count];
            IntPtr read = Marshal.AllocCoTaskMem(4);
            try
            {
                int done = 0;
                while (done < count)
                {
                    byte[] chunk = new byte[count - done];
                    stream.Read(chunk, chunk.Length, read);
                    int n = Marshal.ReadInt32(read);
                    if (n <= 0) throw new EndOfStreamException();
                    Array.Copy(chunk, 0, buffer, done, n);
                    done += n;
                }
            }
            finally { Marshal.FreeCoTaskMem(read); }
            return buffer;
        }

        // ------------------------------------------------------------------
        // IObjectSafety: seguro para dados (as propriedades gravadas são só textos, números e
        // booleanos) e para chamadas. O Office consulta antes de carregar o controle do arquivo.
        // ------------------------------------------------------------------

        const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 1, INTERFACESAFE_FOR_UNTRUSTED_DATA = 2;

        int IObjectSafety.GetInterfaceSafetyOptions(ref Guid iid, out int supported, out int enabled)
        {
            supported = enabled = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            return S_OK;
        }

        int IObjectSafety.SetInterfaceSafetyOptions(ref Guid iid, int mask, int enabled)
        {
            return (mask & ~(INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA)) == 0 ? S_OK : E_FAIL;
        }

        // Valores lidos do arquivo, com o padrão quando faltam
        protected static T Get<T>(IDictionary<string, object> properties, string name, T fallback)
        {
            object value;
            return properties.TryGetValue(name, out value) && value is T ? (T)value : fallback;
        }

        // ------------------------------------------------------------------
        // Win32
        // ------------------------------------------------------------------

        const int SW_HIDE = 0;
        const uint SWP_NOZORDER = 0x4, SWP_NOACTIVATE = 0x10, SWP_SHOWWINDOW = 0x40;
        [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);
        [DllImport("user32.dll")] static extern IntPtr SetParent(IntPtr child, IntPtr parent);
        [DllImport("user32.dll")] static extern IntPtr GetParent(IntPtr h);
    }

    // A janela do controle ativo: filha da janela do UserForm, com o conteúdo do WinForms dentro
    [ComVisible(false)]
    public sealed class OleHost : Control
    {
        const int WS_CHILD = 0x40000000, WS_CLIPSIBLINGS = 0x04000000, WS_CLIPCHILDREN = 0x02000000, WS_POPUP = unchecked((int)0x80000000);
        const int WM_NCHITTEST = 0x84, HTTRANSPARENT = -1;
        readonly OleControl owner;
        internal IntPtr ParentHandle;
        internal bool Transparent;

        internal OleHost(OleControl owner)
        {
            this.owner = owner;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.Parent = ParentHandle;
                cp.Style = (cp.Style & ~WS_POPUP) | WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN;
                return cp;
            }
        }

        protected override void OnPaint(PaintEventArgs e) { owner.PaintHost(e.Graphics, ClientRectangle); }

        protected override void WndProc(ref Message m)
        {
            if (Transparent && m.Msg == WM_NCHITTEST) { m.Result = (IntPtr)HTTRANSPARENT; return; }
            base.WndProc(ref m);
        }
    }

    // ----------------------------------------------------------------------
    // Interfaces do OLE (só o que os controles usam). Internas: o .NET as expõe ao COM pelo
    // QueryInterface, mas elas não entram na biblioteca de tipos.
    // ----------------------------------------------------------------------

    [StructLayout(LayoutKind.Sequential)]
    struct RECT { public int Left, Top, Right, Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    struct SIZEL { public int cx, cy; }

    [StructLayout(LayoutKind.Sequential)]
    struct OLEMSG { public IntPtr hwnd; public int message; public IntPtr wParam, lParam; public int time; public int x, y; }

    [StructLayout(LayoutKind.Sequential)]
    struct OLEINPLACEFRAMEINFO { public int cb; public int fMDIApp; public IntPtr hwndFrame, haccel; public int cAccelEntries; }

    [ComImport, Guid("00000112-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleObject
    {
        [PreserveSig] int SetClientSite(IOleClientSite clientSite);
        [PreserveSig] int GetClientSite(out IOleClientSite clientSite);
        [PreserveSig] int SetHostNames([MarshalAs(UnmanagedType.LPWStr)] string app, [MarshalAs(UnmanagedType.LPWStr)] string obj);
        [PreserveSig] int Close(int saveOption);
        [PreserveSig] int SetMoniker(int which, IntPtr moniker);
        [PreserveSig] int GetMoniker(int assign, int which, out IntPtr moniker);
        [PreserveSig] int InitFromData(IntPtr data, int creation, int reserved);
        [PreserveSig] int GetClipboardData(int reserved, out IntPtr data);
        [PreserveSig] int DoVerb(int verb, IntPtr msg, IOleClientSite activeSite, int index, IntPtr parent, IntPtr posRect);
        [PreserveSig] int EnumVerbs(out IntPtr enumerator);
        [PreserveSig] int Update();
        [PreserveSig] int IsUpToDate();
        [PreserveSig] int GetUserClassID(out Guid clsid);
        [PreserveSig] int GetUserType(int form, [MarshalAs(UnmanagedType.LPWStr)] out string type);
        [PreserveSig] int SetExtent(int aspect, ref SIZEL size);
        [PreserveSig] int GetExtent(int aspect, out SIZEL size);
        [PreserveSig] int Advise(IAdviseSink sink, out int cookie);
        [PreserveSig] int Unadvise(int cookie);
        [PreserveSig] int EnumAdvise(out IntPtr enumerator);
        [PreserveSig] int GetMiscStatus(int aspect, out int status);
        [PreserveSig] int SetColorScheme(IntPtr palette);
    }

    [ComImport, Guid("00000118-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleClientSite
    {
        [PreserveSig] int SaveObject();
        [PreserveSig] int GetMoniker(int assign, int which, out IntPtr moniker);
        [PreserveSig] int GetContainer(out IntPtr container);
        [PreserveSig] int ShowObject();
        [PreserveSig] int OnShowWindow(int show);
        [PreserveSig] int RequestNewObjectLayout();
    }

    [ComImport, Guid("00000119-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleInPlaceSite
    {
        [PreserveSig] int GetWindow(out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp(int enter);
        [PreserveSig] int CanInPlaceActivate();
        [PreserveSig] int OnInPlaceActivate();
        [PreserveSig] int OnUIActivate();
        [PreserveSig] int GetWindowContext(out IOleInPlaceFrame frame, out IOleInPlaceUIWindow document, out RECT pos, out RECT clip,
            ref OLEINPLACEFRAMEINFO info);
        [PreserveSig] int Scroll(SIZEL extent);
        [PreserveSig] int OnUIDeactivate(int undoable);
        [PreserveSig] int OnInPlaceDeactivate();
        [PreserveSig] int DiscardUndoState();
        [PreserveSig] int DeactivateAndUndo();
        [PreserveSig] int OnPosRectChange(ref RECT pos);
    }

    [ComImport, Guid("00000115-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleInPlaceUIWindow
    {
        [PreserveSig] int GetWindow(out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp(int enter);
        [PreserveSig] int GetBorder(out RECT border);
        [PreserveSig] int RequestBorderSpace(ref RECT widths);
        [PreserveSig] int SetBorderSpace(ref RECT widths);
        [PreserveSig] int SetActiveObject(IOleInPlaceActiveObject activeObject, [MarshalAs(UnmanagedType.LPWStr)] string name);
    }

    [ComImport, Guid("00000116-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleInPlaceFrame
    {
        [PreserveSig] int GetWindow(out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp(int enter);
        [PreserveSig] int GetBorder(out RECT border);
        [PreserveSig] int RequestBorderSpace(ref RECT widths);
        [PreserveSig] int SetBorderSpace(ref RECT widths);
        [PreserveSig] int SetActiveObject(IOleInPlaceActiveObject activeObject, [MarshalAs(UnmanagedType.LPWStr)] string name);
        [PreserveSig] int InsertMenus(IntPtr menu, IntPtr widths);
        [PreserveSig] int SetMenu(IntPtr menu, IntPtr olemenu, IntPtr activeObject);
        [PreserveSig] int RemoveMenus(IntPtr menu);
        [PreserveSig] int SetStatusText([MarshalAs(UnmanagedType.LPWStr)] string text);
        [PreserveSig] int EnableModeless(int enable);
        [PreserveSig] int TranslateAccelerator(ref OLEMSG msg, short id);
    }

    [ComImport, Guid("00000113-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleInPlaceObject
    {
        [PreserveSig] int GetWindow(out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp(int enter);
        [PreserveSig] int InPlaceDeactivate();
        [PreserveSig] int UIDeactivate();
        [PreserveSig] int SetObjectRects(ref RECT pos, ref RECT clip);
        [PreserveSig] int ReactivateAndUndo();
    }

    [ComImport, Guid("00000117-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleInPlaceActiveObject
    {
        [PreserveSig] int GetWindow(out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp(int enter);
        [PreserveSig] int TranslateAccelerator(ref OLEMSG msg);
        [PreserveSig] int OnFrameWindowActivate(int activate);
        [PreserveSig] int OnDocWindowActivate(int activate);
        [PreserveSig] int ResizeBorder(ref RECT border, IntPtr window, int frameWindow);
        [PreserveSig] int EnableModeless(int enable);
    }

    [ComImport, Guid("B196B288-BAB4-101A-B69C-00AA00341D07"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleControl
    {
        [PreserveSig] int GetControlInfo(IntPtr info);
        [PreserveSig] int OnMnemonic(ref OLEMSG msg);
        [PreserveSig] int OnAmbientPropertyChange(int dispid);
        [PreserveSig] int FreezeEvents(int freeze);
    }

    [ComImport, Guid("B196B289-BAB4-101A-B69C-00AA00341D07"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleControlSite
    {
        [PreserveSig] int OnControlInfoChanged();
        [PreserveSig] int LockInPlaceActive(int locked);
        [PreserveSig] int GetExtendedControl([MarshalAs(UnmanagedType.IDispatch)] out object control);
        [PreserveSig] int TransformCoords(IntPtr himetric, IntPtr container, int flags);
        [PreserveSig] int TranslateAccelerator(ref OLEMSG msg, int modifiers);
        [PreserveSig] int OnFocus(int gotFocus);
        [PreserveSig] int ShowPropertyFrame();
    }

    [ComImport, Guid("0000010D-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IViewObject
    {
        [PreserveSig] int Draw(int aspect, int index, IntPtr aspectInfo, IntPtr device, IntPtr targetDc, IntPtr hdc, IntPtr bounds,
            IntPtr wbounds, IntPtr cont, IntPtr contArg);
        [PreserveSig] int GetColorSet(int aspect, int index, IntPtr aspectInfo, IntPtr device, IntPtr targetDc, out IntPtr palette);
        [PreserveSig] int Freeze(int aspect, int index, IntPtr aspectInfo, out int freeze);
        [PreserveSig] int Unfreeze(int freeze);
        [PreserveSig] int SetAdvise(int aspects, int advf, IAdviseSink sink);
        [PreserveSig] int GetAdvise(out int aspects, out int advf, out IAdviseSink sink);
    }

    [ComImport, Guid("00000127-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IViewObject2
    {
        [PreserveSig] int Draw(int aspect, int index, IntPtr aspectInfo, IntPtr device, IntPtr targetDc, IntPtr hdc, IntPtr bounds,
            IntPtr wbounds, IntPtr cont, IntPtr contArg);
        [PreserveSig] int GetColorSet(int aspect, int index, IntPtr aspectInfo, IntPtr device, IntPtr targetDc, out IntPtr palette);
        [PreserveSig] int Freeze(int aspect, int index, IntPtr aspectInfo, out int freeze);
        [PreserveSig] int Unfreeze(int freeze);
        [PreserveSig] int SetAdvise(int aspects, int advf, IAdviseSink sink);
        [PreserveSig] int GetAdvise(out int aspects, out int advf, out IAdviseSink sink);
        [PreserveSig] int GetExtent(int aspect, int index, IntPtr device, out SIZEL size);
    }

    [ComImport, Guid("CB5BDC81-93C1-11CF-8F20-00805F2CD064"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IObjectSafety
    {
        [PreserveSig] int GetInterfaceSafetyOptions(ref Guid iid, out int supported, out int enabled);
        [PreserveSig] int SetInterfaceSafetyOptions(ref Guid iid, int mask, int enabled);
    }

    [ComImport, Guid("0000010C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPersist
    {
        [PreserveSig] int GetClassID(out Guid clsid);
    }

    [ComImport, Guid("7FD52380-4E07-101B-AE2D-08002B2EC713"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPersistStreamInit
    {
        [PreserveSig] int GetClassID(out Guid clsid);
        [PreserveSig] int IsDirty();
        [PreserveSig] int Load(IStream stream);
        [PreserveSig] int Save(IStream stream, int clearDirty);
        [PreserveSig] int GetSizeMax(out long size);
        [PreserveSig] int InitNew();
    }

    [ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPersistStream
    {
        [PreserveSig] int GetClassID(out Guid clsid);
        [PreserveSig] int IsDirty();
        [PreserveSig] int Load(IStream stream);
        [PreserveSig] int Save(IStream stream, int clearDirty);
        [PreserveSig] int GetSizeMax(out long size);
    }
}
