using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace SageEditor
{
    // Menu "Sage" na barra de menus do VBE, entre "Janela" e "Ajuda":
    //   Exibir > Explorador e Painel (marcados quando estão à vista; o clique abre ou fecha,
    //            como o Ctrl+R e o Ctrl+J)
    //   Configurações...
    sealed class SageMenu : IDisposable
    {
        const int msoBarTypeMenuBar = 1;
        const int msoControlButton = 1;
        const int msoControlPopup = 10;
        const int msoButtonUp = 0, msoButtonDown = -1;
        const int HelpMenuId = 30010;
        const string Tag = "Sage.Menu";
        const int RefreshInterval = 500; // ms entre as conferências das marcas

        dynamic popup;
        ButtonClick settingsClick;
        readonly List<ButtonClick> clicks = new List<ButtonClick>();
        readonly List<ViewItem> viewItems = new List<ViewItem>();
        int lastRefresh = Environment.TickCount;

        sealed class ViewItem
        {
            public dynamic Button;
            public Func<bool> IsOn;
            public bool? Shown; // a marca que está no menu (null: ainda não conferida)
        }

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

            dynamic viewMenu = popup.Controls.Add(msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
            viewMenu.Caption = Strings.MenuView;
            viewMenu.Tag = "Sage.View";
            object view = viewMenu; // object, para as chamadas abaixo não virarem dinâmicas (com lambdas)
            AddView(view, Strings.MenuExplorer, "Ctrl+R", UnifiedExplorer.Toggle, UnifiedExplorer.IsOpen);
            AddView(view, Strings.MenuPanel, "Ctrl+J", TerminalWindow.Toggle, () => TerminalWindow.IsOpen);

            dynamic settings = popup.Controls.Add(msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            settings.Caption = Strings.MenuSettings;
            settings.Tag = "Sage.Settings";
            settings.Style = 3; // msoButtonIconAndCaption
            settings.FaceId = 548;
            settings.BeginGroup = true;
            settingsClick = new ButtonClick(settings, openSettings);
        }

        void AddView(object view, string caption, string shortcut, Action toggle, Func<bool> isOn)
        {
            dynamic button = ((dynamic)view).Controls.Add(msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            button.Caption = caption;
            button.Style = 2; // msoButtonCaption: a marca aparece no lugar do ícone
            button.ShortcutText = shortcut;
            clicks.Add(new ButtonClick((object)button, delegate { toggle(); lastRefresh = Environment.TickCount - RefreshInterval; }));
            viewItems.Add(new ViewItem { Button = button, IsOn = isOn });
        }

        // Pelo timer do Connect: acerta as marcas do Exibir quando uma janela abre ou fecha
        // (pelo menu, pelos atalhos ou pelo próprio VBE). Só escreve no menu o que mudou.
        public void Poll()
        {
            int now = Environment.TickCount;
            if (unchecked(now - lastRefresh) < RefreshInterval) return;
            lastRefresh = now;
            foreach (ViewItem item in viewItems)
            {
                bool on;
                try { on = item.IsOn(); }
                catch (Exception) { continue; } // VBE ocupado; confere na próxima vez
                if (item.Shown == on) continue;
                item.Button.State = on ? msoButtonDown : msoButtonUp;
                item.Shown = on;
            }
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
            foreach (ButtonClick c in clicks) c.Dispose();
            clicks.Clear();
            viewItems.Clear();
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
