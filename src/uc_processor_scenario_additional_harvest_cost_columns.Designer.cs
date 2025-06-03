namespace FIA_Biosum_Manager
{
    partial class uc_processor_scenario_additional_harvest_cost_columns
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnEditPrev = new System.Windows.Forms.Button();
            this.lblNewlyAdded = new System.Windows.Forms.Label();
            this.lblNewAddedDescription = new System.Windows.Forms.Label();
            this.btnRemoveAll = new System.Windows.Forms.Button();
            this.btnEditNull = new System.Windows.Forms.Button();
            this.btnEditAll = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.lblNullCount = new System.Windows.Forms.Label();
            this.lblDesc = new System.Windows.Forms.Label();
            this.lblType = new System.Windows.Forms.Label();
            this.lblColumnName = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblTitle = new System.Windows.Forms.Label();
            this.uc_processor_scenario_additional_harvest_cost_column_item1 = new FIA_Biosum_Manager.uc_processor_scenario_additional_harvest_cost_column_item();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnEditPrev);
            this.groupBox1.Controls.Add(this.lblNewlyAdded);
            this.groupBox1.Controls.Add(this.lblNewAddedDescription);
            this.groupBox1.Controls.Add(this.btnRemoveAll);
            this.groupBox1.Controls.Add(this.btnEditNull);
            this.groupBox1.Controls.Add(this.btnEditAll);
            this.groupBox1.Controls.Add(this.btnAdd);
            this.groupBox1.Controls.Add(this.lblNullCount);
            this.groupBox1.Controls.Add(this.lblDesc);
            this.groupBox1.Controls.Add(this.lblType);
            this.groupBox1.Controls.Add(this.lblColumnName);
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(895, 398);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
            // 
            // btnEditPrev
            // 
            this.btnEditPrev.ForeColor = System.Drawing.Color.Black;
            this.btnEditPrev.Location = new System.Drawing.Point(618, 367);
            this.btnEditPrev.Name = "btnEditPrev";
            this.btnEditPrev.Size = new System.Drawing.Size(271, 25);
            this.btnEditPrev.TabIndex = 35;
            this.btnEditPrev.Text = "Update All Values Using Previously Entered Data";
            this.btnEditPrev.UseVisualStyleBackColor = true;
            this.btnEditPrev.Click += new System.EventHandler(this.btnEditPrev_Click);
            // 
            // lblNewlyAdded
            // 
            this.lblNewlyAdded.AutoSize = true;
            this.lblNewlyAdded.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNewlyAdded.ForeColor = System.Drawing.Color.Red;
            this.lblNewlyAdded.Location = new System.Drawing.Point(471, 24);
            this.lblNewlyAdded.Name = "lblNewlyAdded";
            this.lblNewlyAdded.Size = new System.Drawing.Size(42, 13);
            this.lblNewlyAdded.TabIndex = 34;
            this.lblNewlyAdded.Text = "00000";
            this.lblNewlyAdded.Visible = false;
            // 
            // lblNewAddedDescription
            // 
            this.lblNewAddedDescription.AutoSize = true;
            this.lblNewAddedDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNewAddedDescription.ForeColor = System.Drawing.Color.Black;
            this.lblNewAddedDescription.Location = new System.Drawing.Point(514, 24);
            this.lblNewAddedDescription.Name = "lblNewAddedDescription";
            this.lblNewAddedDescription.Size = new System.Drawing.Size(305, 13);
            this.lblNewAddedDescription.TabIndex = 33;
            this.lblNewAddedDescription.Text = "new records added to Additional Harvest Costs table";
            this.lblNewAddedDescription.Visible = false;
            // 
            // btnRemoveAll
            // 
            this.btnRemoveAll.ForeColor = System.Drawing.Color.Black;
            this.btnRemoveAll.Location = new System.Drawing.Point(165, 367);
            this.btnRemoveAll.Name = "btnRemoveAll";
            this.btnRemoveAll.Size = new System.Drawing.Size(203, 25);
            this.btnRemoveAll.TabIndex = 32;
            this.btnRemoveAll.Text = "Remove All Scenario Components";
            this.btnRemoveAll.UseVisualStyleBackColor = true;
            this.btnRemoveAll.Click += new System.EventHandler(this.btnRemoveAll_Click);
            // 
            // btnEditNull
            // 
            this.btnEditNull.ForeColor = System.Drawing.Color.Black;
            this.btnEditNull.Location = new System.Drawing.Point(499, 367);
            this.btnEditNull.Name = "btnEditNull";
            this.btnEditNull.Size = new System.Drawing.Size(113, 25);
            this.btnEditNull.TabIndex = 31;
            this.btnEditNull.Text = "Edit Null Values";
            this.btnEditNull.UseVisualStyleBackColor = true;
            this.btnEditNull.Click += new System.EventHandler(this.btnEditNull_Click);
            // 
            // btnEditAll
            // 
            this.btnEditAll.ForeColor = System.Drawing.Color.Black;
            this.btnEditAll.Location = new System.Drawing.Point(393, 367);
            this.btnEditAll.Name = "btnEditAll";
            this.btnEditAll.Size = new System.Drawing.Size(100, 25);
            this.btnEditAll.TabIndex = 30;
            this.btnEditAll.Text = "Edit All Values";
            this.btnEditAll.UseVisualStyleBackColor = true;
            this.btnEditAll.Click += new System.EventHandler(this.btnEditAll_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.ForeColor = System.Drawing.Color.Black;
            this.btnAdd.Location = new System.Drawing.Point(8, 367);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(151, 25);
            this.btnAdd.TabIndex = 29;
            this.btnAdd.Text = "Add Scenario Component";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // lblNullCount
            // 
            this.lblNullCount.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNullCount.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblNullCount.Location = new System.Drawing.Point(355, 53);
            this.lblNullCount.Name = "lblNullCount";
            this.lblNullCount.Size = new System.Drawing.Size(89, 24);
            this.lblNullCount.TabIndex = 28;
            this.lblNullCount.Text = "Null Count";
            // 
            // lblDesc
            // 
            this.lblDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDesc.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblDesc.Location = new System.Drawing.Point(254, 53);
            this.lblDesc.Name = "lblDesc";
            this.lblDesc.Size = new System.Drawing.Size(104, 24);
            this.lblDesc.TabIndex = 28;
            this.lblDesc.Text = "Description";
            // 
            // lblType
            // 
            this.lblType.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblType.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblType.Location = new System.Drawing.Point(153, 48);
            this.lblType.Name = "lblType";
            this.lblType.Size = new System.Drawing.Size(95, 46);
            this.lblType.TabIndex = 28;
            this.lblType.Text = "Scenario /  Treatment";
            // 
            // lblColumnName
            // 
            this.lblColumnName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblColumnName.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblColumnName.Location = new System.Drawing.Point(8, 53);
            this.lblColumnName.Name = "lblColumnName";
            this.lblColumnName.Size = new System.Drawing.Size(139, 24);
            this.lblColumnName.TabIndex = 27;
            this.lblColumnName.Text = "Component Name";
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.uc_processor_scenario_additional_harvest_cost_column_item1);
            this.panel1.Location = new System.Drawing.Point(3, 97);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(889, 264);
            this.panel1.TabIndex = 26;
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(889, 32);
            this.lblTitle.TabIndex = 25;
            this.lblTitle.Text = "Harvest Cost Components";
            // 
            // uc_processor_scenario_additional_harvest_cost_column_item1
            // 
            this.uc_processor_scenario_additional_harvest_cost_column_item1.ColumnName = "";
            this.uc_processor_scenario_additional_harvest_cost_column_item1.DefaultCubicFootDollarValue = "$0.00";
            this.uc_processor_scenario_additional_harvest_cost_column_item1.Description = "";
            this.uc_processor_scenario_additional_harvest_cost_column_item1.ForeColor = System.Drawing.Color.Black;
            this.uc_processor_scenario_additional_harvest_cost_column_item1.Location = new System.Drawing.Point(3, 3);
            this.uc_processor_scenario_additional_harvest_cost_column_item1.Name = "uc_processor_scenario_additional_harvest_cost_column_item1";
            this.uc_processor_scenario_additional_harvest_cost_column_item1.NullCount = "00000";
            this.uc_processor_scenario_additional_harvest_cost_column_item1.ReferenceAdditionalHarvestCostColumnsUserControl = null;
            this.uc_processor_scenario_additional_harvest_cost_column_item1.ReferenceQueries = null;
            this.uc_processor_scenario_additional_harvest_cost_column_item1.Size = new System.Drawing.Size(869, 74);
            this.uc_processor_scenario_additional_harvest_cost_column_item1.TabIndex = 0;
            this.uc_processor_scenario_additional_harvest_cost_column_item1.Type = "";
            // 
            // uc_processor_scenario_additional_harvest_cost_columns
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_scenario_additional_harvest_cost_columns";
            this.Size = new System.Drawing.Size(895, 398);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Panel panel1;
        public System.Windows.Forms.Label lblTitle;
        private uc_processor_scenario_additional_harvest_cost_column_item uc_processor_scenario_additional_harvest_cost_column_item1;
        private System.Windows.Forms.Label lblColumnName;
        private System.Windows.Forms.Label lblNullCount;
        private System.Windows.Forms.Label lblDesc;
        private System.Windows.Forms.Label lblType;
        private System.Windows.Forms.Button btnEditNull;
        private System.Windows.Forms.Button btnEditAll;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnRemoveAll;
        private System.Windows.Forms.Label lblNewAddedDescription;
        private System.Windows.Forms.Label lblNewlyAdded;
        private System.Windows.Forms.Button btnEditPrev;


        
    }
}
