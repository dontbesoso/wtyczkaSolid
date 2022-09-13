﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace SolidEdgeCommunity.AddIn
{
    /// <summary>
    /// Abstract base class for ribbon.
    /// </summary>
    public abstract class Ribbon : IDisposable
    {
        //private int _lastCommandId;
        private Guid _environmentCategory;
        private List<RibbonTab> _tabs = new List<RibbonTab>();

        #region Methods

        /// <summary>
        /// Adds a new ribbon tab.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public RibbonTab AddTab(string name)
        {
            var ribbonTab = new RibbonTab(this, name);
            _tabs.Add(ribbonTab);
            return ribbonTab;
        }

        /// <summary>
        /// Returns a button by id.
        /// </summary>
        /// <param name="commandId"></param>
        /// <returns></returns>
        public RibbonButton GetButton(int commandId)
        {
            return Buttons.Where(x => x.CommandId == commandId).FirstOrDefault();
        }

        /// <summary>
        /// Returns a checkbox by id.
        /// </summary>
        /// <param name="commandId"></param>
        /// <returns></returns>
        public RibbonCheckBox GetCheckBox(int commandId)
        {
            return CheckBoxes.Where(x => x.CommandId == commandId).FirstOrDefault();
        }

        /// <summary>
        /// Returns a RibbonControl by id.
        /// </summary>
        /// <typeparam name="TRibbonControl"></typeparam>
        /// <param name="commandId"></param>
        /// <returns></returns>
        public TRibbonControl GetControl<TRibbonControl>(int commandId) where TRibbonControl : RibbonControl
        {
            return Controls.OfType<TRibbonControl>().Where(x => x.CommandId == commandId).FirstOrDefault();
        }

        /// <summary>
        /// Returns a radio button by id.
        /// </summary>
        /// <param name="commandId"></param>
        /// <returns></returns>
        public RibbonRadioButton GetRadioButton(int commandId)
        {
            return RadioButtons.Where(x => x.CommandId == commandId).FirstOrDefault();
        }

        /// <summary>
        /// Loads ribbon xml from the specified string.
        /// </summary>
        public void LoadXml(string xml)
        {
            _tabs.Clear();

            var xmlns = "http://github.com/SolidEdgeCommunity/SolidEdge/Ribbon";
            var xDocument = XDocument.Parse(xml);
            var xTab = XName.Get("tab", xmlns);
            var xGroup = XName.Get("group", xmlns);

            foreach (var tab in xDocument.Root.Descendants(xTab))
            {
                var tabId = tab.Attribute("name");
                var ribbonTab = AddTab(tabId.Value);

                foreach (var group in tab.Descendants(xGroup))
                {
                    var groupId = group.Attribute("name");
                    var ribbonGroup = ribbonTab.AddGroup(groupId.Value);
                    
                    foreach (var control in group.Descendants())
                    {
                        var controlType = control.Name.LocalName;
                        var controlId = control.Attribute("id");
                        var controlLabel = control.Attribute("label");
                        var controlDropDownGroup = control.Attribute("dropDownGroup");
                        var controlScreentip = control.Attribute("screentip");
                        var controlSupertip = control.Attribute("supertip");
                        var controlImageId = control.Attribute("imageId");
                        var controlEnabled = control.Attribute("enabled");
                        var controlMacro = control.Attribute("macro");
                        var controlMacroParameters = control.Attribute("macroParameters");
                        var controlShowImage = control.Attribute("showImage");
                        var controlShowLabel = control.Attribute("showLabel");
                        var controlWebHelpUrl = control.Attribute("webHelpUrl");

                        RibbonControl ribbonControl = null;
                        int commandId = -1;

                        if (int.TryParse(controlId.Value, out commandId))
                        {
                            if (controlType.Equals("button", StringComparison.OrdinalIgnoreCase))
                            {
                                var ribbonButton = new RibbonButton(commandId);
                                var buttonSize = control.Attribute("size");

                                if (buttonSize != null)
                                {
                                    ribbonButton.TryParseSize(buttonSize.Value);
                                }

                                if (controlDropDownGroup != null)
                                {
                                    ribbonButton.DropDownGroup = controlDropDownGroup.Value;
                                }

                                ribbonControl = ribbonButton;
                            }
                            else if (controlType.Equals("checkBox", StringComparison.OrdinalIgnoreCase))
                            {
                                var ribbonCheckBox = new RibbonCheckBox(commandId);
                                ribbonControl = ribbonCheckBox;
                            }
                            else if (controlType.Equals("radioButton", StringComparison.OrdinalIgnoreCase))
                            {
                                var ribbonRadioButton = new RibbonRadioButton(commandId);
                                ribbonControl = ribbonRadioButton;
                            }

                            if (ribbonControl != null)
                            {
                                //ribbonControl.Name = controlName.Value;
                                ribbonControl.Label = controlLabel.Value;
                                ribbonControl.ScreenTip = controlScreentip.Value;
                                ribbonControl.SuperTip = controlSupertip.Value;

                                if (controlMacro != null)
                                {
                                    ribbonControl.Macro = controlMacro.Value;
                                }

                                if (controlMacroParameters != null)
                                {
                                    ribbonControl.MacroParameters = controlMacroParameters.Value;
                                }

                                if (controlImageId != null)
                                {
                                    ribbonControl.TryParseImageId(controlImageId.Value);
                                }

                                if (controlEnabled != null)
                                {
                                    ribbonControl.TryParseEnabled(controlEnabled.Value);
                                }

                                if (controlShowImage != null)
                                {
                                    ribbonControl.TryParseShowImage(controlShowImage.Value);
                                }

                                if (controlShowLabel != null)
                                {
                                    ribbonControl.TryParseShowLabel(controlShowLabel.Value);
                                }

                                if (controlWebHelpUrl != null)
                                {
                                    ribbonControl.WebHelpURL = controlWebHelpUrl.Value;
                                }

                                ribbonGroup.AddControl(ribbonControl);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Loads ribbon xml from an embedded resource in the specified assembly.
        /// </summary>
        /// <param name="assembly"></param>
        /// <param name="resourceName"></param>
        public void LoadXml(Assembly assembly, string resourceName)
        {
            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream(resourceName)))
            {
                var xml = reader.ReadToEnd();
                LoadXml(xml);
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Returns all RibbonButton controls assigned to the ribbon.
        /// </summary>
        public System.Collections.Generic.IEnumerable<RibbonButton> Buttons
        {
            get
            {
                foreach (var control in this.Controls.OfType<RibbonButton>())
                {
                    yield return control;
                }
            }
        }

        /// <summary>
        /// Returns all RibbonCheckBox controls assigned to the ribbon.
        /// </summary>
        public System.Collections.Generic.IEnumerable<RibbonCheckBox> CheckBoxes
        {
            get
            {
                foreach (var control in this.Controls.OfType<RibbonCheckBox>())
                {
                    yield return control;
                }
            }
        }

        /// <summary>
        /// Returns all controls assigned to the ribbon.
        /// </summary>
        public System.Collections.Generic.IEnumerable<RibbonControl> Controls
        {
            get
            {
                foreach (var tab in Tabs)
                {
                    foreach (var control in tab.Controls)
                    {
                        yield return control;
                    }
                }
            }
        }

        /// <summary>
        /// Returns all RibbonRadioButton controls assigned to the ribbon.
        /// </summary>
        public System.Collections.Generic.IEnumerable<RibbonRadioButton> RadioButtons
        {
            get
            {
                foreach (var control in this.Controls.OfType<RibbonRadioButton>())
                {
                    yield return control;
                }
            }
        }

        public Guid EnvironmentCategory
        {
            get { return _environmentCategory; }
            internal set { _environmentCategory = value; }
        }

        /// <summary>
        /// Returns a RibbonControl by id.
        /// </summary>
        /// <param name="commandId"></param>
        /// <returns></returns>
        public RibbonControl this[int commandId]
        {
            get
            {
                return this.Controls.Where(x => x.CommandId == commandId).FirstOrDefault();
            }
        }

        /// <summary>
        /// Returns an enumerable collection of ribbon tabs.
        /// </summary>
        public IEnumerable<RibbonTab> Tabs { get { return _tabs.AsEnumerable(); } }

        #endregion

        /// <summary>
        /// Called when a user clicks a ribbon control.
        /// </summary>
        /// <param name="control"></param>
        public virtual void OnControlClick(RibbonControl control)
        {
        }

        #region IDisposable implemenation

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
        }

        #endregion

    }
}
