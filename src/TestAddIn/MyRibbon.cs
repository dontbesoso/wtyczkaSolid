﻿using SolidEdgeCommunity.AddIn;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestAddIn
{
    //[RibbonAttribute(SolidEdge.CATID.SEApplication)]
    //[RibbonAttribute(SolidEdge.CATID.SEPart)]
    //[RibbonAttribute(SolidEdge.CATID.SEDMPart)]
    public class MyRibbon : Ribbon
    {
        const string _embeddedResourceName = "TestAddIn.Ribbon.xml";
        private RibbonControl _buttonOpenGlBoxes;
        //private CustomDialog _customDialog;
        private Form1 _myForm;


        public MyRibbon()
            : base()
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();

            this.LoadXml(assembly, "TestAddIn.Ribbon.xml");

            _buttonOpenGlBoxes = GetButton(21);
            _buttonOpenGlBoxes.Click += _buttonOpenGlBoxes_Click;
        }

        public override void OnControlClick(RibbonControl control)
        {
            //var application = MyAddIn.Instance.Application;
            //var documents = application.Documents;
            //var document = (SolidEdgeFramework.SolidEdgeDocument)documents.Add("SolidEdge.PartDocument");
            //document.Close();
        }

        void _buttonOpenGlBoxes_Click(RibbonControl control)
        {
 
            _myForm = new Form1();
            _myForm.ShowDialog();
  
        }
    }
}
