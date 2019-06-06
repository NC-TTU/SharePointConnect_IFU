///<summary>
/// Erstellt von: NC-TTU  
/// Erstellt am: 14.12.17  
/// 
/// Die SharePointFileControlAddin-Klasse ist die Schnittstelle
/// zwischen Assembly und NAV und erlaubt NAV die 
/// SharePointUserControl-Klasse zu nutzen und ihr Daten zu senden.  
/// </summary>

using Microsoft.Dynamics.Framework.UI.Extensibility;
using Microsoft.Dynamics.Framework.UI.Extensibility.WinForms;
using System.Windows.Forms;
using System.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace SharePointConnect
{
    [ControlAddInExport("SharePointConnect")] // Name über den NAV das Client Control Addin anspricht.
    [Serializable]
    public class SharePointFileControlAddin : WinFormsControlAddInBase
    {

        private Panel panel;
        private SharePointUserControl SPUserControl;


        [ApplicationVisible]
        [field: NonSerialized]
        public event MethodInvoker AddInReady = delegate { };

        [ApplicationVisible]
        [field: NonSerialized]
        public event MethodInvoker Create = delegate { };

        [ApplicationVisible]
        [field: NonSerialized]
        public event EventHandler<DragDropEventArgs> Drag_Drop = delegate { };

        protected override Control CreateControl() {


            this.panel = new Panel {
                Dock = DockStyle.Fill,
                Size = new Size(200, 200),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };


            this.SPUserControl = new SharePointUserControl {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Size = panel.ClientSize
            };
            this.panel.Controls.Add(this.SPUserControl);
            this.panel.HandleCreated += (s, e) => AddInReady();
            this.SPUserControl.Create += () => Create();
            this.SPUserControl.Drag_Drop += (s, e) => Drag_Drop(s, e);
            
            return this.panel;
        }

        [ApplicationVisible]
        public void AddSharePointData(System.Collections.ArrayList sharePointData) {
            if (this.SPUserControl != null)
                this.SPUserControl.AddSharePointData(sharePointData);
        }

        [ApplicationVisible]
        public void GetConnection(string baseUrl, string subWebsite, string user, string password, string subPage, string listName, string filter) {
            this.SPUserControl.GetConnection(baseUrl, subWebsite, user, password, subPage, listName, filter);
        }

        [ApplicationVisible]
        public void RefreshSPC() {
            this.SPUserControl.RefreshSPC();
        }

        [ApplicationVisible]
        public void GetConnectionEvent(string baseUrl, string subWebsite, string user, string password, string subPage, string listName, string templateNo, string eventNo) {
            this.SPUserControl.GetConnectionEvent(baseUrl, subWebsite, user, password, subPage, listName, templateNo, eventNo);
        }
    }
}