
namespace CriarXml
{
    partial class Layout
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_Gerar = new System.Windows.Forms.Button();
            this.btn_Selecionar = new System.Windows.Forms.Button();
            this.Seleciona_Arq = new System.Windows.Forms.OpenFileDialog();
            this.txtArquivo = new System.Windows.Forms.TextBox();
            this.instExe = new System.Windows.Forms.TextBox();
            this.btnAbrir = new System.Windows.Forms.Button();
            this.txtCaminho = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btn_Gerar
            // 
            this.btn_Gerar.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btn_Gerar.Location = new System.Drawing.Point(12, 103);
            this.btn_Gerar.Name = "btn_Gerar";
            this.btn_Gerar.Size = new System.Drawing.Size(420, 36);
            this.btn_Gerar.TabIndex = 0;
            this.btn_Gerar.Text = "Gerar";
            this.btn_Gerar.UseVisualStyleBackColor = false;
            this.btn_Gerar.Click += new System.EventHandler(this.button1_Click);
            // 
            // btn_Selecionar
            // 
            this.btn_Selecionar.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btn_Selecionar.Location = new System.Drawing.Point(438, 49);
            this.btn_Selecionar.Name = "btn_Selecionar";
            this.btn_Selecionar.Size = new System.Drawing.Size(132, 36);
            this.btn_Selecionar.TabIndex = 1;
            this.btn_Selecionar.Text = "Selecionar";
            this.btn_Selecionar.UseVisualStyleBackColor = false;
            this.btn_Selecionar.Click += new System.EventHandler(this.btn_Selecionar_Click);
            // 
            // Seleciona_Arq
            // 
            this.Seleciona_Arq.DefaultExt = "*.*xls";
            this.Seleciona_Arq.Filter = "*.*xls|*.*xlsx";
            // 
            // txtArquivo
            // 
            this.txtArquivo.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.txtArquivo.Enabled = false;
            this.txtArquivo.Location = new System.Drawing.Point(12, 56);
            this.txtArquivo.Name = "txtArquivo";
            this.txtArquivo.Size = new System.Drawing.Size(420, 20);
            this.txtArquivo.TabIndex = 2;
            this.txtArquivo.TextChanged += new System.EventHandler(this.txtArquivo_TextChanged);
            // 
            // instExe
            // 
            this.instExe.BackColor = System.Drawing.SystemColors.ControlLight;
            this.instExe.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.instExe.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.instExe.Enabled = false;
            this.instExe.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.instExe.ForeColor = System.Drawing.SystemColors.WindowText;
            this.instExe.Location = new System.Drawing.Point(12, 265);
            this.instExe.Multiline = true;
            this.instExe.Name = "instExe";
            this.instExe.Size = new System.Drawing.Size(558, 153);
            this.instExe.TabIndex = 3;
            this.instExe.Text = "ATENÇÃO! \r\n\r\nOS ARQUIVOS FICARAM SALVO NA PASTA DE EXECUÇÃO DO PROGRAMA";
            this.instExe.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // btnAbrir
            // 
            this.btnAbrir.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnAbrir.Location = new System.Drawing.Point(438, 103);
            this.btnAbrir.Name = "btnAbrir";
            this.btnAbrir.Size = new System.Drawing.Size(132, 36);
            this.btnAbrir.TabIndex = 4;
            this.btnAbrir.Text = "Abrir Pasta";
            this.btnAbrir.UseVisualStyleBackColor = false;
            this.btnAbrir.Click += new System.EventHandler(this.btnAbrir_Click);
            // 
            // txtCaminho
            // 
            this.txtCaminho.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.txtCaminho.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtCaminho.Enabled = false;
            this.txtCaminho.Location = new System.Drawing.Point(12, 173);
            this.txtCaminho.Name = "txtCaminho";
            this.txtCaminho.Size = new System.Drawing.Size(420, 20);
            this.txtCaminho.TabIndex = 5;
            this.txtCaminho.Visible = false;
            // 
            // Layout
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ClientSize = new System.Drawing.Size(598, 357);
            this.Controls.Add(this.txtCaminho);
            this.Controls.Add(this.btnAbrir);
            this.Controls.Add(this.instExe);
            this.Controls.Add(this.txtArquivo);
            this.Controls.Add(this.btn_Selecionar);
            this.Controls.Add(this.btn_Gerar);
            this.Name = "Layout";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Layout_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Gerar;
        private System.Windows.Forms.Button btn_Selecionar;
        private System.Windows.Forms.OpenFileDialog Seleciona_Arq;
        private System.Windows.Forms.TextBox txtArquivo;
        private System.Windows.Forms.TextBox instExe;
        private System.Windows.Forms.Button btnAbrir;
        private System.Windows.Forms.TextBox txtCaminho;
    }
}

