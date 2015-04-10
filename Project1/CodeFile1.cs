using System;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Drawing;
using System.Collections;

using Tekla.Structures;
using Tekla.Structures.Model;

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

            // <profile list>.csv - file location and name
            string csvLocation = "f:/stock list.csv";

            // <profile list>.csv - delimeter
            string delimiterString = ";";

            // list of part names for FL-PL profile check
            string[] partNamesToCheckArray = { "Afstivning", "(Afstivning)", "Vind-X-Plade", "(Vind-X-Plade)", "Løsdele", "(Løsdele)", "Plade", "(Plade)", "Fladstål", "(Fladstål)", "Flange", "(Flange)" };

            // list of part names to include in name AND prefix swaping (should be Plade and Fladstal)
            string[] partNamesToSwapArray = { "Plade", "(Plade)", "Fladstål", "(Fladstål)" };

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // stock list.csv
            //
            // Instructions for preparation:
            // 1. you need original DS stock list,
            // 2. in excel delete all columns but 'Dimension', 'Reserveret' and 'Kvalitet'. This columns should be placed in A, B and C column positions,
            // 3. go through the rows and:
            //       - delete the rows with missing material,
            //       - delete or repair rows with corrupt profile values (look for stuff like: '12x150', '100*5', '15'). Correct formatting is: 'width thickness'.
            // 4. save the file as .csv delimited with semicolon (you can change the delimiter few lines above - delimiterString)
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // preparation of variables                                                                                                                                                                       
            char delimeter = delimiterString[0];

            List<string> partNamesToCheck = new List<string>();
            partNamesToCheck.AddRange(partNamesToCheckArray);

            List<string> partNamesToSwap = new List<string>();
            partNamesToSwap.AddRange(partNamesToSwapArray);

            // Profile list - profiles with attributes (width, thickness, material)
            // if profile is reserved it does not go in this list
            List<List<string>> profileList = new List<List<string>>();
            profileList = csvReader(csvLocation, delimeter);

            // if clause to exit if csvReader didn't succeed
            if ( profileList.Count == 0 ) return;

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
            mainForm form = new mainForm();
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
                var nameOfObject = "";
                var profileOfObject = "";
                var prefixAssemblyOfObject = "";
                var prefixPartOfObject = "";
                bool isFlatProfile = false;

                // get name of the object
                currentObject.GetReportProperty("NAME", ref nameOfObject);

                // get the profile of the object
                currentObject.GetReportProperty("PROFILE", ref profileOfObject);

                // get the prefix of the object
                currentObject.GetReportProperty("ASSEMBLY_DEFAULT_PREFIX", ref prefixAssemblyOfObject);
                currentObject.GetReportProperty("PART_PREFIX", ref prefixPartOfObject);

                // check if profile is flat profile
                if (profileOfObject.StartsWith("FL") || profileOfObject.StartsWith("PL")) isFlatProfile = true;

                // if name is contained in the list of parts to check and profile is a flat profile go in
                if (partNamesToCheck.Contains(nameOfObject) && isFlatProfile)
                {
                    // variables
                    string objectMaterial = "";
                    double objectWidth = -1.0;
                    double objectHeight = -1.0;
                    double objectLength = -1.0;
                    currentObject.GetReportProperty("MATERIAL", ref objectMaterial);
                    currentObject.GetReportProperty("WIDTH", ref objectWidth);
                    currentObject.GetReportProperty("HEIGHT", ref objectHeight);
                    currentObject.GetReportProperty("LENGTH", ref objectLength);

                    // check if profile is in stock list
                    bool inStock = false;
                    inStock = FLCheck(profileList, objectMaterial, objectWidth, objectHeight, objectLength);

                    // check how profile should be changed
                    bool changeToFL = false;
                    bool changeToPL = false;
                    if (inStock && profileOfObject.StartsWith("PL")) changeToFL = true;
                    if (!inStock && profileOfObject.StartsWith("FL")) changeToPL = true;


                    // check how name should be changed
                    bool changeToFladstal = false;
                    bool changeToPlade = false;

                    // this is used to change prefixes
                    bool changeToF = false;     
                    bool changeToC = false;
                    if (partNamesToSwap.Contains(nameOfObject))
                    {
                        if (inStock && nameOfObject.Replace("(", "").Replace(")", "") == "Plade") changeToFladstal = true;
                        if (!inStock && nameOfObject.Replace("(", "").Replace(")", "") == "Fladstål") changeToPlade = true;
                        if (inStock && (prefixPartOfObject != "F" || prefixAssemblyOfObject != "f")) changeToF = true;
                        if (!inStock && (prefixPartOfObject != "C" || prefixAssemblyOfObject != "c")) changeToC = true;
                    }

                    // Functionality for changing the atributes is doubled for beams and plates.
                    // Could this be done in one clause?
                    Beam beam = SelectedObjects.Current as Beam;
                    if (beam != null)
                    {                       
                        if (changeToFL) beam.Profile.ProfileString = "FL" + beam.Profile.ProfileString.ToString().Remove(0, 2);
                        if (changeToPL) beam.Profile.ProfileString = "PL" + beam.Profile.ProfileString.ToString().Remove(0, 2);
                        if (changeToFladstal) beam.Name = "Fladstål";
                        if (changeToF)
                        {
                            beam.AssemblyNumber.Prefix = "f";
                            beam.PartNumber.Prefix = "F";
                        }
                        if (changeToPlade) beam.Name = "Plade";
                        if (changeToC)
                        {
                            beam.AssemblyNumber.Prefix = "c";
                            beam.PartNumber.Prefix = "C";
                        }

                        // add parts to the list of modified parts
                        if (changeToFL || changeToPL || changeToFladstal || changeToPlade || changeToC || changeToF)
                        {
                            partList.Add(beam);
                        }
                    }

                    ContourPlate plate = SelectedObjects.Current as ContourPlate;
                    if (plate != null)
                    {
                        if (changeToFL) plate.Profile.ProfileString = "FL" + plate.Profile.ProfileString.ToString().Remove(0, 2);
                        if (changeToPL) plate.Profile.ProfileString = "PL" + plate.Profile.ProfileString.ToString().Remove(0, 2);
                        if (changeToFladstal) plate.Name = "Fladstål";
                        if (changeToF)
                        {
                            plate.AssemblyNumber.Prefix = "f";
                            plate.PartNumber.Prefix = "F";
                        }
                        if (changeToPlade) plate.Name = "Plade";
                        if (changeToC)
                        {
                            plate.AssemblyNumber.Prefix = "c";
                            plate.PartNumber.Prefix = "C";
                        }

                        // add parts to the list of modified parts
                        if (changeToFL || changeToPL || changeToFladstal || changeToPlade)
                        {
                            partList.Add(plate);
                        }
                    }
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
                DialogResult dialogResult = MessageBox.Show(new Form { TopMost = true }, "Selected objects will be modified.", "FLPL checker", MessageBoxButtons.OKCancel);
                if (dialogResult == DialogResult.OK)
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
                            if (beam != null && selObjEnum.Current.Identifier.ToString() == beam.Identifier.ToString())
                            {
                                if (!beam.Modify()) 
                                {
                                    errCount++;
                                }
                                else 
                                {
                                    modCount++;
                                }
                            }

                            ContourPlate plate = part as ContourPlate;
                            if (plate != null && selObjEnum.Current.Identifier.ToString() == plate.Identifier.ToString())
                            {
                                if (!plate.Modify()) 
                                {
                                    errCount++;
                                }
                                else 
                                {
                                    modCount++;
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
                else if (dialogResult == DialogResult.Cancel)
                {
                    return;
                }
            }
            else
            {
                MessageBox.Show("No parts to modifiy found.", "FLPL checker");
            }
        }

        /// <summary>
        /// reads .csv file and returns a list of profiles with ("width", "thickness", "material"). Skip if profile is reserved.
        /// </summary>
        /// <param name="csvLocation"></param>
        /// <param name="delimeter"></param>
        /// <returns></returns>
        public static List<List<string>> csvReader(string csvLocation, char delimeter)
        {
            List<List<string>> profileList = new List<List<string>>();

            try
            {
                using (StreamReader sr = new StreamReader(csvLocation))
                {
                    int i = 0;
                    while (sr.Peek() >= 0)
                    {
                        String line = sr.ReadLine();

                        // skip first line
                        if (i > 0)
                        {
                            List<string> lineList = new List<string>();

                            List<string> profileData = new List<string>();

                            // function returns width and thickness of the profile
                            profileData = profileCheck(line.Split(delimeter)[0]);

                            // go on only if profileData length is 2 (profile string from .csv is legit)
                            if (profileData.Count == 2)
                            {
                                bool isReserved;

                                // check if this profile is reserved
                                isReserved = reservationCheck(line.Split(delimeter)[1]);

                                if (isReserved == false)
                                {
                                    // check if material string is not empty
                                    if (line.Split(delimeter)[2].Length != 0)
                                    {
                                        lineList.Add(profileData[0]);
                                        lineList.Add(profileData[1]);
                                        lineList.Add(line.Split(delimeter)[2]);

                                        // add to profile attributes to profileList
                                        profileList.Add(lineList);
                                    }
                                    else
                                    {
                                        MessageBox.Show("Illegitimate profile line\n\nMaterial in line: \n" + (i + 1).ToString(), "FLPL checker");
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Illegitimate profile line\n\nProfile in line: \n" + (i + 1).ToString(), "FLPL checker");
                                break;
                            }
                        }

                        i += 1;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Could not find stock list in location:\n" + csvLocation + "\n-----------------------------------------------------------------------------------\n" + e.ToString(), "FLPL checker");
            }

            return profileList;
        }

        /// <summary>
        /// check profile string from .csv if it is legit and return width and thickness
        /// </summary>
        /// <param name="profile"></param>
        /// <returns></returns> returns a list (width, thickness)
        public static List<string> profileCheck(string profile)
        {
            // set delimeter for the profile string
            string delimeterString = " ";
            char delimeter = delimeterString[0];

            // declare list and variables
            List<string> profileData = new List<string>();
            string width = "";
            string thickness = "";

            // if string has delimeter inside take it, if not ignore this .csv line
            if (profile.Split(delimeter).Length == 2)
            {
                width = profile.Split(delimeter)[0];
                thickness = profile.Split(delimeter)[1];

                // add profile data to list
                profileData.Add(width);
                profileData.Add(thickness);
            }
            return profileData;
        }

        /// <summary>
        /// check reserved string from .csv if it is not empty
        /// </summary>
        /// <param name="profile"></param>
        /// <returns></returns> returns a bool - false if not reserved
        public static bool reservationCheck(string reservation)
        {
            bool isReserved = false;

            // if string is not empty we should ignore this .csv line
            if (reservation.Length != 0)
            {
                isReserved = true;
            }
            return isReserved;
        }

        /// <summary>
        /// checks if profile is in profile list. If yes, returns bool=true
        /// </summary>
        /// <param name="objectMaterial"></param>
        /// <param name="objectWidth"></param>
        /// <param name="objectHeight"></param>
        /// <param name="objectLength"></param>
        /// <returns></returns>
        public static bool FLCheck(List<List<string>> profileList, string objectMaterial, double objectWidth, double objectHeight, double objectLength)
        {
            bool isFL = false;

            // round doubles to whole values
            objectWidth = Math.Round(objectWidth, 0, MidpointRounding.AwayFromZero);
            objectHeight = Math.Round(objectHeight, 0, MidpointRounding.AwayFromZero);
            objectLength = Math.Round(objectLength, 0, MidpointRounding.AwayFromZero);

            foreach (List<string> profileLine in profileList)
            {
                // check if material of current object starts with material string from .csv (e.g.: "S275JR", "S275")
                if (objectMaterial.StartsWith(profileLine[2]))
                {
                    // check for matching thickness
                    if (profileLine[1] == objectWidth.ToString())
                    {
                        // check if 'Length' of 'Height' of current object matches width from profile list ('Length' and 'Height' as Tekla report properties)
                        if (profileLine[0] == objectLength.ToString() || profileLine[0] == objectHeight.ToString())
                        {
                            isFL = true;
                            break;
                        }
                    }
                }
            }
            return isFL;
        }
    }

    /// <summary>
    /// Main form
    /// All, selected and close buttons
    /// </summary>
    public class mainForm : Form
    {
        private System.Windows.Forms.Label labelMyForm;
        private System.Windows.Forms.Button buttonAll;
        private System.Windows.Forms.Button buttonSelected;
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

        private void buttonAll_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Yes;
        }

        private void buttonSelected_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.No;
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        public mainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonAll = new System.Windows.Forms.Button();
            this.buttonSelected = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            
            // buttonAll
            this.buttonAll.Location = new System.Drawing.Point(13, 38);
            this.buttonAll.Name = "buttonAll";
            this.buttonAll.Size = new System.Drawing.Size(75, 23);
            this.buttonAll.TabIndex = 0;
            this.buttonAll.Text = "All";
            this.buttonAll.UseVisualStyleBackColor = true;
            this.buttonAll.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.buttonAll.Click += new System.EventHandler(this.buttonAll_Click);

            // buttonSelected
            this.buttonSelected.Location = new System.Drawing.Point(100, 38);
            this.buttonSelected.Name = "buttonSelected";
            this.buttonSelected.Size = new System.Drawing.Size(75, 23);
            this.buttonSelected.TabIndex = 0;
            this.buttonSelected.Text = "Selected";
            this.buttonSelected.UseVisualStyleBackColor = true;
            this.buttonSelected.DialogResult = System.Windows.Forms.DialogResult.No;
            this.buttonSelected.Click += new System.EventHandler(this.buttonSelected_Click);

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
            this.label1.Text = "Check FL-PL for:";
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Size = new System.Drawing.Size(200, 20);

            // label2
            this.label2.Text = "V 1.0 / 10.4.2015 / zagar.domen@gmail.com";
            this.label2.Location = new System.Drawing.Point(13, 73);
            this.label2.Size = new System.Drawing.Size(250, 20);

            // Form1
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Name = "mainForm";
            this.Text = "FLPL checker";

            // holds top position
            this.TopMost = true;
            
            this.ClientSize = new System.Drawing.Size(278, 100);
            this.Controls.Add(this.buttonAll);
            this.Controls.Add(this.buttonSelected);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.ResumeLayout(false);
        }
    }
}