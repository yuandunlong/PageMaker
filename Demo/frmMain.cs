using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using WinWordControl;
//using Microsoft.Office.Interop.Word;

namespace WordInDOTNET
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class frmMain : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button button1;
		private WinWordControl.WinWordControl objWinWordControl;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtClientName;
		private System.Windows.Forms.TextBox txtRecNo;
		private System.Windows.Forms.Button btnClientName;
		private System.Windows.Forms.Button btnRecNo;
		private System.Windows.Forms.Button btnCurrentDate;
		private System.Windows.Forms.Button btnMarkError;
		private System.Windows.Forms.Button btnRemoveError;
		private System.Windows.Forms.Button btnSaveDocument;
		private System.Windows.Forms.Button btnRemoveBookMarks;
		private System.Windows.Forms.Button btnPrintView;
		private System.Windows.Forms.Button btnNormalView;
		private System.Windows.Forms.Button btnWebView;
		private System.Windows.Forms.Button btnShowMenuBar;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private Button button2;
        int i1;
        string aa;
        private Button button3;
        private Button bt_name;
        private Button bt_Locat;
        private Office.CommandBarButton maincombar;

        //maincombar



		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmMain()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.button1 = new System.Windows.Forms.Button();
            this.objWinWordControl = new WinWordControl.WinWordControl();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtClientName = new System.Windows.Forms.TextBox();
            this.txtRecNo = new System.Windows.Forms.TextBox();
            this.btnClientName = new System.Windows.Forms.Button();
            this.btnRecNo = new System.Windows.Forms.Button();
            this.btnCurrentDate = new System.Windows.Forms.Button();
            this.btnMarkError = new System.Windows.Forms.Button();
            this.btnRemoveError = new System.Windows.Forms.Button();
            this.btnSaveDocument = new System.Windows.Forms.Button();
            this.btnRemoveBookMarks = new System.Windows.Forms.Button();
            this.btnPrintView = new System.Windows.Forms.Button();
            this.btnNormalView = new System.Windows.Forms.Button();
            this.btnWebView = new System.Windows.Forms.Button();
            this.btnShowMenuBar = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.bt_name = new System.Windows.Forms.Button();
            this.bt_Locat = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(19, 9);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(115, 24);
            this.button1.TabIndex = 1;
            this.button1.Text = "Load Document";
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // objWinWordControl
            // 
            this.objWinWordControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.objWinWordControl.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.objWinWordControl.Location = new System.Drawing.Point(306, 0);
            this.objWinWordControl.Name = "objWinWordControl";
            this.objWinWordControl.Range = null;
            this.objWinWordControl.Size = new System.Drawing.Size(406, 663);
            this.objWinWordControl.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(10, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 17);
            this.label1.TabIndex = 4;
            this.label1.Text = "Client\'s Name";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(10, 112);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(105, 17);
            this.label2.TabIndex = 5;
            this.label2.Text = "Record Number";
            // 
            // txtClientName
            // 
            this.txtClientName.Location = new System.Drawing.Point(19, 69);
            this.txtClientName.Name = "txtClientName";
            this.txtClientName.Size = new System.Drawing.Size(120, 21);
            this.txtClientName.TabIndex = 6;
            this.txtClientName.Text = "Some Name";
            // 
            // txtRecNo
            // 
            this.txtRecNo.Location = new System.Drawing.Point(19, 129);
            this.txtRecNo.Name = "txtRecNo";
            this.txtRecNo.Size = new System.Drawing.Size(120, 21);
            this.txtRecNo.TabIndex = 7;
            this.txtRecNo.Text = "456789123";
            // 
            // btnClientName
            // 
            this.btnClientName.Location = new System.Drawing.Point(144, 69);
            this.btnClientName.Name = "btnClientName";
            this.btnClientName.Size = new System.Drawing.Size(38, 17);
            this.btnClientName.TabIndex = 9;
            this.btnClientName.Text = "-->";
            this.btnClientName.Click += new System.EventHandler(this.btnClientName_Click);
            // 
            // btnRecNo
            // 
            this.btnRecNo.Location = new System.Drawing.Point(144, 129);
            this.btnRecNo.Name = "btnRecNo";
            this.btnRecNo.Size = new System.Drawing.Size(38, 17);
            this.btnRecNo.TabIndex = 10;
            this.btnRecNo.Text = "-->";
            this.btnRecNo.Click += new System.EventHandler(this.btnRecNo_Click);
            // 
            // btnCurrentDate
            // 
            this.btnCurrentDate.Location = new System.Drawing.Point(10, 172);
            this.btnCurrentDate.Name = "btnCurrentDate";
            this.btnCurrentDate.Size = new System.Drawing.Size(153, 26);
            this.btnCurrentDate.TabIndex = 11;
            this.btnCurrentDate.Text = "Insert Current Date";
            this.btnCurrentDate.Click += new System.EventHandler(this.btnCurrentDate_Click);
            // 
            // btnMarkError
            // 
            this.btnMarkError.Location = new System.Drawing.Point(10, 214);
            this.btnMarkError.Name = "btnMarkError";
            this.btnMarkError.Size = new System.Drawing.Size(90, 25);
            this.btnMarkError.TabIndex = 12;
            this.btnMarkError.Text = "Mark Error";
            this.btnMarkError.Click += new System.EventHandler(this.btnMarkError_Click);
            // 
            // btnRemoveError
            // 
            this.btnRemoveError.Location = new System.Drawing.Point(10, 257);
            this.btnRemoveError.Name = "btnRemoveError";
            this.btnRemoveError.Size = new System.Drawing.Size(105, 25);
            this.btnRemoveError.TabIndex = 15;
            this.btnRemoveError.Text = "RemoveHighlighting";
            this.btnRemoveError.Click += new System.EventHandler(this.btnRemoveError_Click);
            // 
            // btnSaveDocument
            // 
            this.btnSaveDocument.Location = new System.Drawing.Point(10, 344);
            this.btnSaveDocument.Name = "btnSaveDocument";
            this.btnSaveDocument.Size = new System.Drawing.Size(134, 24);
            this.btnSaveDocument.TabIndex = 16;
            this.btnSaveDocument.Text = "Save Document";
            this.btnSaveDocument.Click += new System.EventHandler(this.btnSaveDocument_Click);
            // 
            // btnRemoveBookMarks
            // 
            this.btnRemoveBookMarks.Location = new System.Drawing.Point(10, 301);
            this.btnRemoveBookMarks.Name = "btnRemoveBookMarks";
            this.btnRemoveBookMarks.Size = new System.Drawing.Size(134, 24);
            this.btnRemoveBookMarks.TabIndex = 17;
            this.btnRemoveBookMarks.Text = "Remove bookmarks";
            this.btnRemoveBookMarks.Click += new System.EventHandler(this.btnRemoveBookMarks_Click);
            // 
            // btnPrintView
            // 
            this.btnPrintView.Location = new System.Drawing.Point(10, 395);
            this.btnPrintView.Name = "btnPrintView";
            this.btnPrintView.Size = new System.Drawing.Size(144, 25);
            this.btnPrintView.TabIndex = 18;
            this.btnPrintView.Text = "Print Layout View";
            this.btnPrintView.Click += new System.EventHandler(this.btnPrintView_Click);
            // 
            // btnNormalView
            // 
            this.btnNormalView.Location = new System.Drawing.Point(10, 430);
            this.btnNormalView.Name = "btnNormalView";
            this.btnNormalView.Size = new System.Drawing.Size(144, 25);
            this.btnNormalView.TabIndex = 19;
            this.btnNormalView.Text = "Normal View";
            this.btnNormalView.Click += new System.EventHandler(this.btnNormalView_Click);
            // 
            // btnWebView
            // 
            this.btnWebView.Location = new System.Drawing.Point(10, 464);
            this.btnWebView.Name = "btnWebView";
            this.btnWebView.Size = new System.Drawing.Size(90, 25);
            this.btnWebView.TabIndex = 20;
            this.btnWebView.Text = "Web View";
            this.btnWebView.Click += new System.EventHandler(this.btnWebView_Click);
            // 
            // btnShowMenuBar
            // 
            this.btnShowMenuBar.Location = new System.Drawing.Point(10, 507);
            this.btnShowMenuBar.Name = "btnShowMenuBar";
            this.btnShowMenuBar.Size = new System.Drawing.Size(90, 25);
            this.btnShowMenuBar.TabIndex = 21;
            this.btnShowMenuBar.Text = "Show MenuBar";
            this.btnShowMenuBar.Click += new System.EventHandler(this.btnShowMenuBar_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(10, 548);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 22;
            this.button2.Text = "测试";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(10, 593);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 23;
            this.button3.Text = "关闭";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // bt_name
            // 
            this.bt_name.Location = new System.Drawing.Point(124, 465);
            this.bt_name.Name = "bt_name";
            this.bt_name.Size = new System.Drawing.Size(75, 23);
            this.bt_name.TabIndex = 24;
            this.bt_name.Text = "Name";
            this.bt_name.UseVisualStyleBackColor = true;
            this.bt_name.Click += new System.EventHandler(this.bt_name_Click);
            // 
            // bt_Locat
            // 
            this.bt_Locat.Location = new System.Drawing.Point(124, 509);
            this.bt_Locat.Name = "bt_Locat";
            this.bt_Locat.Size = new System.Drawing.Size(75, 23);
            this.bt_Locat.TabIndex = 25;
            this.bt_Locat.Text = "Locat";
            this.bt_Locat.UseVisualStyleBackColor = true;
            this.bt_Locat.Click += new System.EventHandler(this.bt_Locat_Click);
            // 
            // frmMain
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 14);
            this.ClientSize = new System.Drawing.Size(712, 663);
            this.Controls.Add(this.bt_Locat);
            this.Controls.Add(this.bt_name);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btnShowMenuBar);
            this.Controls.Add(this.btnWebView);
            this.Controls.Add(this.btnNormalView);
            this.Controls.Add(this.btnPrintView);
            this.Controls.Add(this.btnRemoveBookMarks);
            this.Controls.Add(this.btnSaveDocument);
            this.Controls.Add(this.btnRemoveError);
            this.Controls.Add(this.btnMarkError);
            this.Controls.Add(this.btnCurrentDate);
            this.Controls.Add(this.btnRecNo);
            this.Controls.Add(this.btnClientName);
            this.Controls.Add(this.txtRecNo);
            this.Controls.Add(this.txtClientName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.objWinWordControl);
            this.Controls.Add(this.button1);
            this.Name = "frmMain";
            this.Text = "Main";
            this.Closing += new System.ComponentModel.CancelEventHandler(this.frmMain_Closing);
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			System.Windows.Forms.Application.Run(new frmMain());
		}

		private void frmMain_Load(object sender, System.EventArgs e)
		{
            objWinWordControl.LoadDocument(@"d:\1.doc");
            objWinWordControl.document.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;

		}

		/// <summary>
		/// Loads the document. If it is already loaded, it will first unload and load again.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void button1_Click(object sender, System.EventArgs e)
		{

			string filNm;
			openFileDialog1.Multiselect = false;
			openFileDialog1.Filter = "MS-Word Files (*.doc,*.dot) | *.doc;*.dot";
			if(openFileDialog1.ShowDialog() == DialogResult.OK)
			{
            filNm = openFileDialog1.FileName;
			}
			else return;
			//MessageBox.Show("Please wait while the document is being displayed");
			try
			{
				objWinWordControl.CloseControl();

			}
			catch{}
			finally
			{ 
				objWinWordControl.document=null;
				WinWordControl.WinWordControl.wd=null;
				WinWordControl.WinWordControl.wordWnd=0;
			}
			try
			{

				//Load the template used for testing.
				objWinWordControl.LoadDocument(filNm);
			}
			catch(Exception ex){String err = ex.Message;}
			btnClientName.Enabled=true;
			btnRecNo.Enabled=true;
			btnCurrentDate.Enabled=true;
		}

		private void frmMain_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			objWinWordControl.RestoreCommandBars();
			objWinWordControl.CloseControl();
		}

		
		/// <summary>
		/// Searches the bookmark by name and places text at the text
		/// </summary>
		/// <param name="BookMarkName"></param>
		/// <param name="BookMarkText"></param>
 
		private void WriteToBookMark(string BookMarkName, string BookMarkText)
		{
			try
			{
				Word.Document wd = objWinWordControl.document;
				Word.Application wa = wd.Application;
				int bookmark_cnt = wd.Bookmarks.Count;
				int i;
				for(i=1;i<=bookmark_cnt;i++)
				{
					object o = (object)i;
					if(BookMarkName.ToLower().Trim() == wd.Bookmarks.Item(ref o).Name.ToLower().Trim())
					{
						wd.Bookmarks.Item(ref o).Select();
						wa.Selection.TypeText(BookMarkText);
					}
				}
			}
			catch(Exception ex)
			{
				String err = ex.Message;
			}
		}


		/// <summary>
		/// Custom: Writes the client name at the bookmark.
		/// This template is pre-defined with 3 bookmarks.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnClientName_Click(object sender, System.EventArgs e)
		{
			WriteToBookMark("bmrkClientName",txtClientName.Text);
			//btnClientName.Enabled=false;
		}

		/// <summary>
		/// Custom: Writes the client name at the bookmark.
		/// This template is pre-defined with 3 bookmarks.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnRecNo_Click(object sender, System.EventArgs e)
		{
			WriteToBookMark("bmrkRecNo",txtRecNo.Text);
			//btnRecNo.Enabled=false;
		}

		/// <summary>
		/// Custom: Writes the client name at the bookmark.
		/// This template is pre-defined with 3 bookmarks.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnCurrentDate_Click(object sender, System.EventArgs e)
		{
			WriteToBookMark("bmrkCurrentDate",DateTime.Now.ToString());
			//btnCurrentDate.Enabled=false;
		}


		/// <summary>
		/// Mark the error by selecting some text
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnMarkError_Click(object sender, System.EventArgs e)
		{
			try
			{
				string strText = objWinWordControl.document.Application.Selection.Text;
				frmMarkError f= new frmMarkError();
				f.txtChanged.Text = strText;
				f.txtOriginal.Text = strText;
				DialogResult dr= f.ShowDialog();

				if(dr==DialogResult.OK)
				{
					objWinWordControl.document.Application.Selection.Text=f.txtChanged.Text;
					objWinWordControl.document.Application.Selection.FormattedText.HighlightColorIndex=Word.WdColorIndex.wdYellow;
					string bkmrkname = "bkmrk_err_" + DateTime.Now.Day.ToString().PadLeft(2,'0') + DateTime.Now.Month.ToString().PadLeft(2,'0')+ DateTime.Now.Year.ToString()+ DateTime.Now.Hour.ToString().PadLeft(2,'0')+ DateTime.Now.Minute.ToString().PadLeft(2,'0')+ DateTime.Now.Second.ToString().PadLeft(2,'0')+ DateTime.Now.Ticks.ToString().PadLeft(8,'0');
					object o = objWinWordControl.document.Application.Selection.Range;
					objWinWordControl.document.Bookmarks.Add(bkmrkname,ref o);
					objWinWordControl.document.Application.Selection.FormattedText.HighlightColorIndex=Word.WdColorIndex.wdYellow;
				}

				// Do something else like making entry to database.
			}
			catch{}
		}

		/// <summary>
		/// Remove the Error bookmarks
		/// This might be required if the document is being sent for further quality check.
		/// In that case, the errors marked by current QA must not be shown to next QA
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnRemoveError_Click(object sender, System.EventArgs e)
		{
			try
			{
				string BookMarkName="";
				Word.Document wd = objWinWordControl.document;
				Word.Application wa = wd.Application;
				int bookmark_cnt = wd.Bookmarks.Count;
				int i;
				for(i=1;i<=bookmark_cnt;i++)
				{
					object o = (object)i;
					BookMarkName=wd.Bookmarks.Item(ref o).Name;
					if(BookMarkName.Substring(0,10)=="bkmrk_err_")
					{
						wd.Bookmarks.Item(ref o).Select();
						wa.Selection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
					}
				}

				//Do something else
			}
			catch(Exception ex)
			{
				String err = ex.Message;
			}		
		}

		/// <summary>
		/// Save the document. Calls the Word's save method
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSaveDocument_Click(object sender, System.EventArgs e)
		{

            try
			{

                this.objWinWordControl.document.Save();
			}
			catch
			{}
		}

		/// <summary>
		/// Clear all the bookmarks. 
		/// This may be required before submitting the transcript to the client.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnRemoveBookMarks_Click(object sender, System.EventArgs e)
		{
			try
			{
				Word.Document wd = objWinWordControl.document;
				Word.Application wa = wd.Application;
				int bookmark_cnt = wd.Bookmarks.Count;
				int i;

				for(i=1;i<=bookmark_cnt;i++)
				{
					object o = (object)wd.Bookmarks.Count;
					wd.Bookmarks.Item(ref o).Delete();
				}

				MessageBox.Show("Bookmarks removed");
			}
			catch(Exception ex)
			{
				String err = ex.Message;
				MessageBox.Show("Error occured while removing bookmarks");
			}		
		}

		/// <summary>
		/// Change the view
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnPrintView_Click(object sender, System.EventArgs e)
		{
			try
			{
				objWinWordControl.document.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
			}
			catch{}
		}


		/// <summary>
		/// Change the view
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnNormalView_Click(object sender, System.EventArgs e)
		{
			try
			{
				objWinWordControl.document.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
			}
			catch{}
		}


		/// <summary>
		/// Change the view
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnWebView_Click(object sender, System.EventArgs e)
		{
			try
			{
				objWinWordControl.document.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdWebView;
			}
			catch{}
		}

		/// <summary>
		/// If you want to show the menubar to the user. 
		/// (Useful in cases where too many functionalities of word are being used)
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnShowMenuBar_Click(object sender, System.EventArgs e)
		{
			try
			{
				objWinWordControl.document.ActiveWindow.Application.CommandBars["Menu Bar"].Enabled=true;
			}
			catch{}
		}

        private void button2_Click(object sender, EventArgs e)
        {
            
            
//private  static  wdCharacter = 1;
//Const wdExtent = 1;
//Const wdExtend = 1;
//Const wdGoToBookmark = -1;
//Const wdLine = 5;
//Const wdCell = 12;
//Const wdTableFormatSimple2 = 2;
//Const wdAlignParagraphRight = 2;
//Const wdYellow = 7;
//Const wdToggle = 9999998;
//Const wdAlignParagraphCenter = 1;
//Const wdSentence = 3;
//Const wdParagraph = 4;
//Const wdStory = 6;
//Const wdMove = 0;


            try
            {


                Word.Document wd = objWinWordControl.document;
                Word.Application wa1 = wd.Application;
                Word.Range r1 = objWinWordControl.Range;

                int i11;
                string bb;

                //wd.Paragraphs.Add();
                //  Word.Application oWord;


                ////x选择集操作 r1.Paragraphs.Item(1).Range, Item中的数字只能是1，表示当前段落；
                //wa.Selection.Font.NameFarEast = "隶书";
                ////wa.Selection.WholeStory();//文档全选
                //r1 = wa.Selection.Range;

                //r1 = r1.Paragraphs.Item(1).Range;
                //aa = r1.Text;
                //MessageBox.Show(aa);

                   /// 书签操作
                   string bookmark_11;
                   int i12;
                   string time1 = DateTime.Now.ToString();
                for (i12 = 1; i12 <= wd.Paragraphs.Count; i12++)
                {
                    object o = wd.Paragraphs.Item(i12).Range;
                     bookmark_11 = "abcd"+i12.ToString();
                     //MessageBox.Show(bookmark_11);

                     objWinWordControl.document.Bookmarks.Add(bookmark_11, ref o);

                   //  MessageBox.Show(bookmark_11);
                     i1 = i12;

                   }
                string time2= DateTime.Now.ToString();
                MessageBox.Show(i1.ToString());
                object o111 = (object)i1;
                MessageBox.Show(objWinWordControl.document.Bookmarks.Item(o111).Range.Text);
                MessageBox.Show(time1+"----------"+time2);

                // objWinWordControl.document.Bookmarks.DefaultSorting = Word.WdBookmarkSortBy.wdSortByLocation;//这类赋值需带上word.WdBookmarkSortBy.

                //string aaa= objWinWordControl.document.Bookmarks.Item(11).Name;
                // MessageBox.Show(aaa);





                ////段落操作
                //for (i11 = 1; i11 <= 10; i11++)
                //{

                ////aa = wd.Paragraphs.Item(1).Range.End.ToString();
            //    r1 = wd.Paragraphs.Item(2).Range;
            //    aa = r1.Text;
                //r1.Font.NameFarEast = "隶书";
                //r1.Font.Name = "Times New Roman";
                //r1.Paragraphs.CharacterUnitFirstLineIndent = 5;




               // MessageBox.Show(aa);
//



                //wa.Selection.MoveDown(Word.WdParagraphAlignment,3,1);// unit:=wdParagraph, Extend:=wdExtend 
                //wa.Selection.MoveStart(1);
                // wa.Selection.MoveEnd();


                // Selection.MoveUp unit:=wdParagraph
                //Selection.MoveDown unit:=wdParagraph, Extend:=wdExtend


                //禁用保存和另存的代码如下:
//Sub 禁用保存()
//Application.CommandBars("File").Controls(4).Enabled = False
//Application.CommandBars("File").Controls(5).Enabled = False
//End Sub
//启用保存和另存的代码如下:
//Sub 启用保存()
//Application.CommandBars("File").Controls(4).Enabled = True
//Application.CommandBars("File").Controls(5).Enabled = True
//End Sub
                //Word.Document wa11= objWinWordControl.document;
                //Word.Application waaa = wa11.Application;

                //由当前selection或Paragraphs的选择集（rang)获取BookmarkID，在由BookmarkID获得当前标签名称或相关内容；
              
            //    Word.Selection slct1 = waaa.Selection;
            //    Word.Paragraph opp; Word.Range rr1;
            //    Word.Bookmark bk132;
            //    int no111;
                               

               
            //    string bmk2 = slct1.BookmarkID.ToString();
            //   // MessageBox.Show(bmk2);
            //    int noid =System.Int32.Parse( bmk2);
            //    string srr1 = slct1.Bookmarks.Item(1).Range.Text;
            //   // MessageBox.Show(srr1);
            //    rr1 = slct1.Bookmarks.Item(1).Range;
            //    string bmkname = rr1.Bookmarks.Item(1).Name;
            //   // MessageBox.Show(bmkname);
            //   // MessageBox.Show(rr1.Text);
            //   // WordDoc.Bookmarks.get_Item(BookmarkID).Range.Start;
            //    //Object oRngoBookMarkStart = wa11.Bookmarks.Item(bmk2).Range.Start;
            //    bk132=slct1.Bookmarks.Item(1);
            //    string hgff = bk132.Name;
                
            //    MessageBox.Show("书签No." + bmkname+"\r" + bk132.Range.Text);
            //    //MessageBox.Show(bmk2);
            //    object o = (object)bmk2;
            //    MessageBox.Show(o.ToString());
            //    string srr12 = bk132.Name;
            //    //MessageBox.Show(o.ToString());
            //    bk132.Range.Font.NameFarEast = "隶书";
            //    bk132.Range.Font.Size = 28;
            //    bk132.Range.Font.Bold = 12;


            //    MessageBox.Show(srr12+"0000000000000000000000000000000000000");
            //    no111 = noid;

            //   //// Word.Application bb = waaa.CommandBars.ActiveMenuBar;
            //   // Office.CommandBar bbwww= waaa.CommandBars.ActiveMenuBar;
            //   // Word.Paragraph opp;Word.Range rr1;

            //   // object oMissing = System.Reflection.Missing.Value;
            //   // object oEndOfDoc = "endofdoc"; /**//* endofdoc is a predefined bookmark */
            //   // wa11.Bookmarks.DefaultSorting = Word.WdBookmarkSortBy.wdSortByLocation;
            //   // rr1 = wa11.Bookmarks.Item(3).Range;//由当前selection或Paragraphs的选择集（rang)获取BookmarkID，在由BookmarkID获得当前标签名称或相关内容；
            //   // opp = wa11.Content.Paragraphs.Item(100);
            //   // opp.Range.InsertParagraphAfter();
            //   // string bmkname =rr1.Bookmarks.Item(1).Name;//Item(1)的参数为1，表示当前书签；
            //   // string bmk1 = rr1.BookmarkID.ToString();
            //   // MessageBox.Show(bmk1);
            //   // MessageBox.Show(bmkname);


            //    ///


            //    object o1 = (object)no111;
            //    MessageBox.Show(o1.ToString());
            //    string text1 = objWinWordControl.document.Bookmarks.Item(ref o1).Range.Text;
            //    MessageBox.Show(text1);


            }

            catch
            { }







        }

                #region 替换

        /////<summary>
        ///// 在word 中查找一个字符串直接替换所需要的文本
        ///// </summary>
        ///// <param name="strOldText">原文本</param>
        ///// <param name="strNewText">新文本</param>
        ///// <returns></returns>
        ///// 
                
        //public bool Replace(string strOldText, string strNewText)
        //{

        //    Word.Document oDoc = objWinWordControl.document;
        //        Word.Application wa1 = oDoc.Application;
        //       // oDoc = objWinWordControl.document;
        //        object missing = System.Reflection.Missing.Value;//重要


        //    if (oDoc == null)
      
        //    oDoc.Content.Find.Text = strOldText;
        //    object FindText, ReplaceWith, Replace;// 
        //    FindText = strOldText;//要查找的文本
        //    ReplaceWith = strNewText;//替换文本
        //    Replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;/**//*wdReplaceAll - 替换找到的所有项。
        //                                              * wdReplaceNone - 不替换找到的任何项。
        //                                            * wdReplaceOne - 替换找到的第一项。
        //                                            * */
        //    oDoc.Content.Find.ClearFormatting();//移除Find的搜索文本和段落格式设置
        //    if (oDoc.Content.Find.Execute( ref FindText, ref missing,
        //        ref missing, ref missing,
        //        ref missing, ref missing,
        //        ref missing, ref missing, 
        //        ref missing, ref ReplaceWith, 
        //        ref Replace, ref missing, 
        //        ref missing, ref missing, ref missing))
        //    {
        //        return true;
        //    }
        //    return false;
        //}

        //public bool SearchReplace(string strOldText, string strNewText)
        //{

        //    Word.Document wd = objWinWordControl.document;
        //    Word.Application oWordApplic = wd.Application;
        //    // oDoc = objWinWordControl.document;
        //    object missing = System.Reflection.Missing.Value;//重要
        //    object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;

        //    //首先清除任何现有的格式设置选项，然后设置搜索字符串 strOldText。
        //    oWordApplic.Selection.Find.ClearFormatting();
        //    oWordApplic.Selection.Find.Text = strOldText;

        //    oWordApplic.Selection.Find.Replacement.ClearFormatting();
        //    oWordApplic.Selection.Find.Replacement.Text = strNewText;

        //    if (oWordApplic.Selection.Find.Execute(
        //        ref missing, ref missing, ref missing, ref missing, ref missing,
        //        ref missing, ref missing, ref missing, ref missing, ref missing,
        //        ref replaceAll, ref missing, ref missing, ref missing, ref missing))
        //    {
        //        return true;
        //    }
        //    return false;
        //}

        #endregion


            
            





        

        private void button3_Click(object sender, EventArgs e)
        {
            objWinWordControl.document.Saved = true;
            objWinWordControl.document.Close();
            objWinWordControl.CloseControl();
            objWinWordControl.Dispose();
            this.Close();
        }

        private void bt_name_Click(object sender, EventArgs e)
        {
            /// 书签操作
            Word.Document wd = objWinWordControl.document;
            Word.Application wa1 = wd.Application;

            try
            {

                string bookmark_11;
                int i12;

                for (i12 = 1; i12 <= 20; i12++)
                {

                    object o = wd.Paragraphs.Item(i12).Range;
                    bookmark_11 ="1-"+ i12.ToString();
                    wd.Bookmarks.Add(bookmark_11, ref o);
                    MessageBox.Show(bookmark_11);

                }
            }
            catch
            { };
        
        }
        private void bt_Locat_Click(object sender, EventArgs e)
        {

            objWinWordControl.document.Bookmarks.DefaultSorting = Word.WdBookmarkSortBy.wdSortByLocation;
        }

       
	}	
	
}