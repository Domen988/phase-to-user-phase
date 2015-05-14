using System;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Drawing;
using System.Collections;

using Tekla.Structures;
using Tekla.Structures.Model;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// This macro is used to copy the comment of object's phase to user phase attribute of an object.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Ideas for improvements
// - it's kind of silly that we read the comment of the phase for the user phase, 
//     + we could use a custom phase property for setting the user phase number
//     + check out: http://teklastructures.support.tekla.com/200/en/mod_custom_phase_properties
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

namespace Tekla.Technology.Akit.UserScript
{     
    public class Script
    {
        // to work in Visual Studio uncomment next line:
		public static void Main()
        // to use this code as Tekla macro uncomment next line:
        //public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////            
            // Settings

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            Model Model = new Model();

            // select object types for selector 
            System.Type[] Types = new System.Type[2];
            Types.SetValue(typeof(Beam), 0);
            Types.SetValue(typeof(ContourPlate), 1);

            // instantiate model object enumerator before if clauses
            Tekla.Structures.Model.ModelObjectEnumerator SelectedObjects = Model.GetModelObjectSelector().GetAllObjectsWithType(ModelObject.ModelObjectEnum.UNKNOWN);
            // =======================================================================================
            // dialog for object selection
            DialogResult dr = new DialogResult();
            mainForm form = new mainForm("Set user phase for:", "All", "Selected");
            dr = form.ShowDialog();
            if (dr == DialogResult.Yes)     // 'Yes' is used for all objects
            {
                SelectedObjects = Model.GetModelObjectSelector().GetAllObjectsWithType(Types);
            }
            else if (dr == DialogResult.No) // 'No' is used to get selected objects
            {
                SelectedObjects = new Tekla.Structures.Model.UI.ModelObjectSelector().GetSelectedObjects();
            }
            else 
            {
               return;
            }
            // =======================================================================================

            // list of changed objects
            ArrayList partList = new ArrayList();

            while (SelectedObjects.MoveNext())
            {
                var currentObject = SelectedObjects.Current;
                var currObjPhaseComment = "";
                var currObjUserPhase = "";
                
                Phase currObjPhase = new Phase();

                currentObject.GetPhase(out currObjPhase);
                // phase comment gets copied to user phase
                currObjPhaseComment = currObjPhase.PhaseComment;
                currentObject.GetUserProperty("USER_PHASE", ref currObjUserPhase);

                if (currObjUserPhase != currObjPhaseComment)
                {
                    //currentObject.SetUserProperty("USER_PHASE", currObjPhaseComment);
                    partList.Add(currentObject);
                }
            }
            
            // select objects that are in list for modification
            Tekla.Structures.Model.UI.ModelObjectSelector mos = new Tekla.Structures.Model.UI.ModelObjectSelector();
            mos.Select(partList);

            // modified object count
            var modCount = 0;
            var errCount = 0;

            // exit if there is no parts to modify
            if (partList.Count != 0)
            {
                // confirm modification
                DialogResult drConfirmation = new DialogResult();
                mainForm formConfirmation = new mainForm("Selected objects will be modified", "Refresh", "Ok");
                drConfirmation = formConfirmation.ShowDialog();
                if (drConfirmation == DialogResult.Yes)     // 'Yes' is used to refresh selection
                {
                    mos.Select(partList);
                }
                else if (drConfirmation == DialogResult.No) // 'No' is used to confirm 
                {
                    // if OK, then go through list and modify each part
                    Tekla.Structures.Model.ModelObjectEnumerator selObjEnum = Model.GetModelObjectSelector().GetAllObjectsWithType(ModelObject.ModelObjectEnum.CONTOURPLATE);
                    selObjEnum = new Tekla.Structures.Model.UI.ModelObjectSelector().GetSelectedObjects();

                    // modify only objects that are in part list for modification and in current selection
                    while (selObjEnum.MoveNext())
                    {
                        foreach (var part in partList)
                        {
                            Beam beam = part as Beam;
                            ContourPlate plate = part as ContourPlate;
                            if (beam != null && selObjEnum.Current.Identifier.ToString() == beam.Identifier.ToString() || plate != null && selObjEnum.Current.Identifier.ToString() == plate.Identifier.ToString())
                            {
                                try
                                {
                                    var currentObject = selObjEnum.Current;
                                    var currObjPhaseComment = "";
                                    var currObjUserPhase = "";
                                    Phase currObjPhase = new Phase();

                                    currentObject.GetPhase(out currObjPhase);
                                    // phase comment gets copied to user phase
                                    currObjPhaseComment = currObjPhase.PhaseComment;
                                    currentObject.GetUserProperty("USER_PHASE", ref currObjUserPhase);

                                    if (currObjUserPhase != currObjPhaseComment)
                                    {
                                        currentObject.SetUserProperty("USER_PHASE", currObjPhaseComment);
                                    }
                                    modCount++;
                                }
                                catch 
                                {
                                    errCount++;
                                }
                            }
                        }
                    }
                    if (errCount != 0)
                    {
                        MessageBox.Show("Warning\n# of objects which didn't modify:\n" + errCount + "\n\n# of changed objects:\n" + modCount, "FLPL checker");
                    }
                    else
                    {
                        MessageBox.Show("# of changed objects:\n" + modCount, "FLPL checker");
                    }
                }
                else
                {
                    return;
                }
            }
            else
            {
                MessageBox.Show("No parts to modifiy found.", Globals.appName);
            }
        }
    }

    public static class Globals
    {
        public const String appName = "Phase to user phase"; // Modifiable in Code
        public const String versionDate = "V 1.0 / 29.4.2015 / zagar.domen@gmail.com";
    }

    /// <summary>
    /// Main form
    /// All, selected and close buttons
    /// </summary>
    public class mainForm : Form
    {
        private System.Windows.Forms.Label labelMyForm;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;

        private void MyMessageForm_Load(object sender, EventArgs e)
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.MinimizeBox = false;
            this.labelMyForm = new System.Windows.Forms.Label();
            this.labelMyForm.Text = "Are you sure you want to… ?";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Yes;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.No;
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.None;
        }

        // mainForm is created with three parameters: question text, button 1 text, button 2 text
        // third button is always cancel
        public mainForm(string formMessage, string button1text, string button2text)
        {
            InitializeComponent(formMessage, button1text, button2text);
        }

        private void InitializeComponent(string formMessage, string button1text, string button2text)
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            
            // button1
            this.button1.Location = new System.Drawing.Point(13, 38);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = button1text;
            this.button1.UseVisualStyleBackColor = true;
           /*
            if (button1text == "Refresh")
            {
                this.button1.DialogResult = System.Windows.Forms.DialogResult.None;
                this.button1.Click += new System.EventHandler(this.button3_Click);
            }
            else
            */
            {
                this.button1.DialogResult = System.Windows.Forms.DialogResult.Yes;
                this.button1.Click += new System.EventHandler(this.button1_Click);
            }
           // this.button1.Click += new System.EventHandler(this.button1_Click);

            // button2
            this.button2.Location = new System.Drawing.Point(100, 38);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 0;
            this.button2.Text = button2text;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.DialogResult = System.Windows.Forms.DialogResult.No;
            this.button2.Click += new System.EventHandler(this.button2_Click);

            // buttonCancel
            this.buttonCancel.Location = new System.Drawing.Point(187, 38);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 0;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);

            // label1
            this.label1.Text = formMessage;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Size = new System.Drawing.Size(200, 20);

            // label2
            this.label2.Text = Globals.versionDate;
            this.label2.Location = new System.Drawing.Point(13, 73);
            this.label2.Size = new System.Drawing.Size(250, 20);

            // Form1
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Name = "mainForm";
            this.Text = Globals.appName;

            // holds top position
            this.TopMost = true;
            
            this.ClientSize = new System.Drawing.Size(278, 100);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.ResumeLayout(false);
        }
    }
}