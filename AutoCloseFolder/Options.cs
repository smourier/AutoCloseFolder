using System;
using System.ComponentModel;
using Microsoft.VisualStudio.Shell;

namespace AutoCloseFolder
{
    public class Options : DialogPage
    {
        [Category("Solution Explorer")]
        [DisplayName("Periodically Collapse Files and Folders")]
        [Description("Periodically collapse files and folder nodes in Solution Explorer.")]
        [DefaultValue(true)]
        public bool CollapseFolders { get; set; } = true;

        [Category("Solution Explorer")]
        [DisplayName("Collapsing Period")]
        [Description("Collapsing period in milliseconds.")]
        [DefaultValue(false)]
        public int Period { get; set; } = _defaultPeriod;

        [Category("Solution Explorer")]
        [DisplayName("Collapsing Period After Node Expansion")]
        [Description("Collapsing period after any node expansion in milliseconds.")]
        [DefaultValue(false)]
        public int PeriodOnExpansion { get; set; } = _defaultPeriodOnExpansion;

        [Category("Solution Explorer")]
        [DisplayName("Collapse Files and Folders On Close")]
        [Description("Collapse nodes in Solution Explorer on close.")]
        [DefaultValue(true)]
        public bool CollapseOnClose { get; set; } = true;

        [Category("Solution Explorer")]
        [DisplayName("Collapse Solution Folders On Close")]
        [Description("Collapse solution folders in Solution Explorer when collapsing on close.")]
        [DefaultValue(true)]
        public bool CollapseSolutionFolders { get; set; } = true;

        [Category("Solution Explorer")]
        [DisplayName("Collapse Projects On Close")]
        [Description("Collapse projects in Solution Explorer when collapsing on close.")]
        [DefaultValue(false)]
        public bool CollapseProjects { get; set; }

        private const int _defaultPeriodOnExpansion = 120000;
        private const int _defaultPeriod = 30000;

        internal int FinalPeriod => Math.Max(_defaultPeriod, Period);
        internal int FinalPeriodOnExpansion => Math.Max(_defaultPeriodOnExpansion, PeriodOnExpansion);
    }
}
