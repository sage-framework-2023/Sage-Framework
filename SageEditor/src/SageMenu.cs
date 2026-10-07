using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace SageEditor
{
    // Menu "Sage" na barra de menus do VBE, entre "Janela" e "Ajuda".
    sealed class SageMenu : IDisposable
    {
        const int msoBarTypeMenuBar = 1;
        const int msoControlButton = 1;
        const int msoControlPopup = 10;
        const int HelpMenuId = 30010;
        const string Tag = "Sage.Menu";

        dynamic popup;
        ButtonClick settingsClick;

        public SageMenu(dynamic vbe, Action openSettings)
        {
            dynamic bar = null;
            foreach (dynamic cb in vbe.CommandBars)
                if (cb.Type == msoBarTypeMenuBar) { bar = cb; break; }
            if (bar == null) throw new InvalidOperationException("Barra de menus do VBE não encontrada.");

            RemoveLeftovers(vbe);

            object before = Type.Missing;
            for (int i = 1; i <= bar.Controls.Count; i++)
                if (bar.Controls[i].Id == HelpMenuId) { before = i; break; }

            popup = bar.Controls.Add(msoControlPopup, Type.Missing, Type.Missing, before, true);
            popup.Caption = "Sa&ge";
            popup.Tag = Tag;

            dynamic settings = popup.Controls.Add(msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            settings.Caption = Strings.MenuSettings;
            settings.Tag = "Sage.Settings";
            settings.Style = 3; // msoButtonIconAndCaption
            settings.FaceId = 548;
            settingsClick = new ButtonClick(settings, openSettings);
        }

        static void RemoveLeftovers(dynamic vbe)
        {
            for (int guard = 0; guard < 10; guard++)
            {
                dynamic old = vbe.CommandBars.FindControl(Type.Missing, Type.Missing, Tag);
                if (old == null) break;
                old.Delete();
            }
        }

        public void Dispose()
        {
            if (settingsClick != null) { settingsClick.Dispose(); settingsClick = null; }
            if (popup != null)
            {
                try { popup.Delete(); } catch (COMException) { }
                popup = null;
            }
        }
    }

    // Evento Click de um CommandBarButton do Office (_CommandBarButtonEvents).
    [ComVisible(true), Guid("000C0351-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ICommandBarButtonEvents
    {
        [DispId(1)]
        void Click([In, MarshalAs(UnmanagedType.IDispatch)] object ctrl, [In, Out] ref bool cancelDefault);
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public sealed class ButtonClick : ICommandBarButtonEvents, IDisposable
    {
        readonly Action action;
        readonly bool cancel;
        IConnectionPoint point;
        int cookie;

        internal ButtonClick(object button, Action action) : this(button, action, false) { }

        // cancel: o comando original do VBE não roda (o Sage o substitui)
        internal ButtonClick(object button, Action action, bool cancel)
        {
            this.action = action;
            this.cancel = cancel;
            Guid iid = typeof(ICommandBarButtonEvents).GUID;
            ((IConnectionPointContainer)button).FindConnectionPoint(ref iid, out point);
            point.Advise(this, out cookie);
        }

        public void Click(object ctrl, ref bool cancelDefault)
        {
            try { action(); }
            catch (Exception ex) { Log.Error(ex); }
            if (cancel) cancelDefault = true;
        }

        public void Dispose()
        {
            if (point == null) return;
            try { point.Unadvise(cookie); } catch (COMException) { }
            point = null;
        }
    }
}
