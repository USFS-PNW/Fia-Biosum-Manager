namespace FIA_Biosum_Manager
{
    partial class uc_processor_scenario_additional_harvest_cost_column_item
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
            this.txtColumnName = new System.Windows.Forms.TextBox();
            this.txtRxScenario = new System.Windows.Forms.TextBox();
            this.lblNullValueCount = new System.Windows.Forms.Label();
            this.grpColumnName = new System.Windows.Forms.GroupBox();
            this.btnColumnNameRemove = new System.Windows.Forms.Button();
            this.btnColumnNameEdit = new System.Windows.Forms.Button();
            this.grpColumnValues = new System.Windows.Forms.GroupBox();
            this.btnGo = new System.Windows.Forms.Button();
            this.cmbEdit = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtCubicFootDollarValue = new System.Windows.Forms.TextBox();
            this.txtDesc = new System.Windows.Forms.TextBox();
            this.grpColumnName.SuspendLayout();
            this.grpColumnValues.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtColumnName
            // 
            this.txtColumnName.Location = new System.Drawing.Point(13, 33);
            this.txtColumnName.Margin = new System.Windows.Forms.Padding(4);
            this.txtColumnName.Name = "txtColumnName";
            this.txtColumnName.Size = new System.Drawing.Size(184, 22);
            this.txtColumnName.TabIndex = 0;
            this.txtColumnName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtColumnName_KeyPress);
            // 
            // txtRxScenario
            // 
            this.txtRxScenario.Location = new System.Drawing.Point(207, 33);
            this.txtRxScenario.Margin = new System.Windows.Forms.Padding(4);
            this.txtRxScenario.Name = "txtRxScenario";
            this.txtRxScenario.Size = new System.Drawing.Size(97, 22);
            this.txtRxScenario.TabIndex = 1;
            this.txtRxScenario.TextChanged += new System.EventHandler(this.txtRxScenario_TextChanged);
            this.txtRxScenario.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRxScenario_KeyPress);
            // 
            // lblNullValueCount
            // 
            this.lblNullValueCount.AutoSize = true;
            this.lblNullValueCount.ForeColor = System.Drawing.Color.Black;
            this.lblNullValueCount.Location = new System.Drawing.Point(504, 33);
            this.lblNullValueCount.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblNullValueCount.Name = "lblNullValueCount";
            this.lblNullValueCount.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.lblNullValueCount.Size = new System.Drawing.Size(48, 17);
            this.lblNullValueCount.TabIndex = 2;
            this.lblNullValueCount.Text = "00000";
            // 
            // grpColumnName
            // 
            this.grpColumnName.Controls.Add(this.btnColumnNameRemove);
            this.grpColumnName.Controls.Add(this.btnColumnNameEdit);
            this.grpColumnName.ForeColor = System.Drawing.Color.Black;
            this.grpColumnName.Location = new System.Drawing.Point(561, 5);
            this.grpColumnName.Margin = new System.Windows.Forms.Padding(4);
            this.grpColumnName.Name = "grpColumnName";
            this.grpColumnName.Padding = new System.Windows.Forms.Padding(4);
            this.grpColumnName.Size = new System.Drawing.Size(147, 80);
            this.grpColumnName.TabIndex = 3;
            this.grpColumnName.TabStop = false;
            // 
            // btnColumnNameRemove
            // 
            this.btnColumnNameRemove.Location = new System.Drawing.Point(32, 46);
            this.btnColumnNameRemove.Margin = new System.Windows.Forms.Padding(4);
            this.btnColumnNameRemove.Name = "btnColumnNameRemove";
            this.btnColumnNameRemove.Size = new System.Drawing.Size(87, 27);
            this.btnColumnNameRemove.TabIndex = 1;
            this.btnColumnNameRemove.Text = "Remove";
            this.btnColumnNameRemove.UseVisualStyleBackColor = true;
            this.btnColumnNameRemove.Click += new System.EventHandler(this.btnColumnNameRemove_Click);
            // 
            // btnColumnNameEdit
            // 
            this.btnColumnNameEdit.Location = new System.Drawing.Point(32, 17);
            this.btnColumnNameEdit.Margin = new System.Windows.Forms.Padding(4);
            this.btnColumnNameEdit.Name = "btnColumnNameEdit";
            this.btnColumnNameEdit.Size = new System.Drawing.Size(87, 27);
            this.btnColumnNameEdit.TabIndex = 0;
            this.btnColumnNameEdit.Text = "Edit";
            this.btnColumnNameEdit.UseVisualStyleBackColor = true;
            this.btnColumnNameEdit.Click += new System.EventHandler(this.btnColumnNameEdit_Click);
            // 
            // grpColumnValues
            // 
            this.grpColumnValues.Controls.Add(this.btnGo);
            this.grpColumnValues.Controls.Add(this.cmbEdit);
            this.grpColumnValues.Controls.Add(this.label1);
            this.grpColumnValues.Controls.Add(this.txtCubicFootDollarValue);
            this.grpColumnValues.ForeColor = System.Drawing.Color.Black;
            this.grpColumnValues.Location = new System.Drawing.Point(716, 2);
            this.grpColumnValues.Margin = new System.Windows.Forms.Padding(4);
            this.grpColumnValues.Name = "grpColumnValues";
            this.grpColumnValues.Padding = new System.Windows.Forms.Padding(4);
            this.grpColumnValues.Size = new System.Drawing.Size(405, 80);
            this.grpColumnValues.TabIndex = 4;
            this.grpColumnValues.TabStop = false;
            this.grpColumnValues.Text = "Assign Values";
            // 
            // btnGo
            // 
            this.btnGo.Location = new System.Drawing.Point(335, 44);
            this.btnGo.Margin = new System.Windows.Forms.Padding(4);
            this.btnGo.Name = "btnGo";
            this.btnGo.Size = new System.Drawing.Size(63, 26);
            this.btnGo.TabIndex = 7;
            this.btnGo.Text = "Go";
            this.btnGo.UseVisualStyleBackColor = true;
            this.btnGo.Click += new System.EventHandler(this.btnGo_Click);
            // 
            // cmbEdit
            // 
            this.cmbEdit.FormattingEnabled = true;
            this.cmbEdit.Items.AddRange(new object[] {
            "Assign default value to all component occurrences",
            "Assign default value to all component NULL values",
            "Assign previously entered values ",
            "Edit all values",
            "Edit NULL values"});
            this.cmbEdit.Location = new System.Drawing.Point(12, 46);
            this.cmbEdit.Margin = new System.Windows.Forms.Padding(4);
            this.cmbEdit.Name = "cmbEdit";
            this.cmbEdit.Size = new System.Drawing.Size(313, 24);
            this.cmbEdit.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 20);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 17);
            this.label1.TabIndex = 5;
            this.label1.Text = "Default $/ac:";
            // 
            // txtCubicFootDollarValue
            // 
            this.txtCubicFootDollarValue.Location = new System.Drawing.Point(119, 17);
            this.txtCubicFootDollarValue.Margin = new System.Windows.Forms.Padding(4);
            this.txtCubicFootDollarValue.Name = "txtCubicFootDollarValue";
            this.txtCubicFootDollarValue.Size = new System.Drawing.Size(80, 22);
            this.txtCubicFootDollarValue.TabIndex = 4;
            this.txtCubicFootDollarValue.Text = "$0.00";
            this.txtCubicFootDollarValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtCubicFootDollarValue.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCubicFootDollarValue_KeyPress);
            this.txtCubicFootDollarValue.Leave += new System.EventHandler(this.txtCubicFootDollarValue_Leave);
            // 
            // txtDesc
            // 
            this.txtDesc.Location = new System.Drawing.Point(313, 5);
            this.txtDesc.Margin = new System.Windows.Forms.Padding(4);
            this.txtDesc.Multiline = true;
            this.txtDesc.Name = "txtDesc";
            this.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtDesc.Size = new System.Drawing.Size(181, 79);
            this.txtDesc.TabIndex = 5;
            this.txtDesc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDesc_KeyPress);
            // 
            // uc_processor_scenario_additional_harvest_cost_column_item
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.txtDesc);
            this.Controls.Add(this.grpColumnValues);
            this.Controls.Add(this.grpColumnName);
            this.Controls.Add(this.lblNullValueCount);
            this.Controls.Add(this.txtRxScenario);
            this.Controls.Add(this.txtColumnName);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "uc_processor_scenario_additional_harvest_cost_column_item";
            this.Size = new System.Drawing.Size(1133, 90);
            this.grpColumnName.ResumeLayout(false);
            this.grpColumnValues.ResumeLayout(false);
            this.grpColumnValues.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtColumnName;
        private System.Windows.Forms.TextBox txtRxScenario;
        private System.Windows.Forms.Label lblNullValueCount;
        private System.Windows.Forms.GroupBox grpColumnName;
        private System.Windows.Forms.Button btnColumnNameEdit;
        private System.Windows.Forms.GroupBox grpColumnValues;
        private System.Windows.Forms.Button btnColumnNameRemove;
        private System.Windows.Forms.TextBox txtDesc;
        private System.Windows.Forms.TextBox txtCubicFootDollarValue;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnGo;
        private System.Windows.Forms.ComboBox cmbEdit;
    }
}
