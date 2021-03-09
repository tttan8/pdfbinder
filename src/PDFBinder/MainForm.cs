using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace PDFBinder
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            UpdateUI();
        }

        public void AddInputFile(string file)
        {
            switch (Combiner.TestSourceFile(file))
            {
                case Combiner.SourceTestResult.Unreadable:
                    MessageBox.Show(string.Format("File could not be opened as a PDF document:\n\n{0}", file), "Illegal file type", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case Combiner.SourceTestResult.Protected:
                    MessageBox.Show(string.Format("PDF document does not allow copying:\n\n{0}", file), "Permission denied", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    break;
                case Combiner.SourceTestResult.Ok:
                    inputListBox.Items.Add(file);
                    break;
            }
        }

        public void UpdateUI()
        {
            if (inputListBox.Items.Count < 2)
            {
                completeButton.Enabled = false;
                helpLabel.Text = "Drop PDF-documents in the box above, or choose \"add document\" from the toolbar";
            }
            else
            {
                completeButton.Enabled = true;
                helpLabel.Text = "Click the \"bind!\" button when you are done adding documents";
            }

            if (inputListBox.SelectedIndex < 0)
            {
                removeButton.Enabled = moveUpButton.Enabled = moveDownButton.Enabled = false;
            }
            else
            {
                removeButton.Enabled = true;
                moveUpButton.Enabled = (inputListBox.SelectedIndex > 0);
                moveDownButton.Enabled = (inputListBox.SelectedIndex < inputListBox.Items.Count - 1);
            }
        }

        private void inputListBox_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop, false) ? DragDropEffects.All : DragDropEffects.None;
        }

        private void inputListBox_DragDrop(object sender, DragEventArgs e)
        {
            var fileNames = (string[]) e.Data.GetData(DataFormats.FileDrop);
            Array.Sort(fileNames);

            foreach (var file in fileNames)
            {
                AddInputFile(file);
            }

            UpdateUI();
        }

        private void combineButton_Click(object sender, EventArgs e)
        {
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                using (var combiner = new Combiner(saveFileDialog.FileName))
                {
                    progressBar.Visible = true;
                    this.Enabled = false;

                    for (int i = 0; i < inputListBox.Items.Count; i++)
                    {
                        combiner.AddFile((string)inputListBox.Items[i],i+1);
                        progressBar.Value = (int)(((i + 1) / (double)inputListBox.Items.Count) * 100);
                    }


                    this.Enabled = true;
                    progressBar.Visible = false;
                }

                System.Diagnostics.Process.Start(saveFileDialog.FileName);
            }
        }

        private void addFileButton_Click(object sender, EventArgs e)
        {
            if (addFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string file in addFileDialog.FileNames)
                {
                    AddInputFile(file);
                }

                UpdateUI();
            }
        }

        private void inputListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateUI();
        }

        private void removeButton_Click(object sender, EventArgs e)
        {
            inputListBox.Items.Remove(inputListBox.SelectedItem);
        }

        private void moveItemButton_Click(object sender, EventArgs e)
        {
            object dataItem = inputListBox.SelectedItem;
            int index = inputListBox.SelectedIndex;

            if (sender == moveUpButton)
                index--;
            if (sender == moveDownButton)
                index++;

            inputListBox.Items.Remove(dataItem);
            inputListBox.Items.Insert(index, dataItem);

            inputListBox.SelectedIndex = index;
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            // 来自：https://blog.csdn.net/jsjpanxiaoyu/article/details/52922119
            //Document:（文档）生成pdf必备的一个对象,生成一个Document示例
            Document document = new Document(PageSize.A4, 30, 30, 5, 5);
            //为该Document创建一个Writer实例：
            PdfWriter.GetInstance(document, new FileStream("e://" + "Chap0101.pdf", FileMode.Create));
            //打开当前Document
            document.Open();
            //为当前Document添加内容：
            document.Add(new Paragraph("Hello World"));
            //另起一行。有几种办法建立一个段落，如：
            Paragraph p1 = new Paragraph(new Chunk("This is my first paragraph.\n", FontFactory.GetFont(FontFactory.HELVETICA, 12)));
            Paragraph p2 = new Paragraph(new Phrase("This is my second paragraph.", FontFactory.GetFont(FontFactory.HELVETICA, 12)));
            Paragraph p3 = new Paragraph("This is my third paragraph.", FontFactory.GetFont(FontFactory.HELVETICA, 12));
            //所有有些对象将被添加到段落中：
            p1.Add("you can add string here\n\t");
            p1.Add(new Chunk("you can add chunks \n")); p1.Add(new Phrase("or you can add phrases.\n"));
            document.Add(p1); document.Add(p2); document.Add(p3);

            //创建了一个内容为“hello World”、红色、斜体、COURIER字体、尺寸20的一个块：
            Chunk chunk = new Chunk("Hello world", FontFactory.GetFont(FontFactory.COURIER, 20, iTextSharp.text.Font.COURIER, new iTextSharp.text.Color(255, 0, 0)));
            document.Add(chunk);
            //如果你希望一些块有下划线或删除线，你可以通过改变字体风格简单做到：
            Chunk chunk1 = new Chunk("This text is underlined", FontFactory.GetFont(FontFactory.HELVETICA, 12, iTextSharp.text.Font.UNDEFINED));
            Chunk chunk2 = new Chunk("This font is of type ITALIC | STRIKETHRU", FontFactory.GetFont(FontFactory.HELVETICA, 12, iTextSharp.text.Font.ITALIC | iTextSharp.text.Font.STRIKETHRU));
            //改变块的背景
            chunk2.SetBackground(new iTextSharp.text.Color(0xFF, 0xFF, 0x00));
            //上标/下标
            chunk1.SetTextRise(5);
            document.Add(chunk1);
            document.Add(chunk2);

            //外部链接示例：
            Anchor anchor = new Anchor("website", FontFactory.GetFont(FontFactory.HELVETICA, 12, iTextSharp.text.Font.UNDEFINED, new iTextSharp.text.Color(0, 0, 255)));

            anchor.Reference = "http://itextsharp.sourceforge.net/";
            anchor.Name = "website";
            //内部链接示例：
            Anchor anchor1 = new Anchor("This is an internal link\n\n");
            anchor1.Name = "link1";
            Anchor anchor2 = new Anchor("Click here to jump to the internal link\n\f");
            anchor2.Reference = "#link1";
            document.Add(anchor); document.Add(anchor1); document.Add(anchor2);

            //排序列表示例：
            List list = new List(true, 20);
            list.Add(new iTextSharp.text.ListItem("First line"));
            list.Add(new iTextSharp.text.ListItem("The second line is longer to see what happens once the end of the line is reached. Will it start on a new line?"));
            list.Add(new iTextSharp.text.ListItem("Third line"));
            document.Add(list);

            //文本注释：
            Annotation a = new Annotation("authors", "Maybe its because I wanted to be an author myself that I wrote iText.");
            document.Add(a);

            //包含页码没有任何边框的页脚。
            HeaderFooter footer = new HeaderFooter(new Phrase("This is page: "), true);
            footer.Border = iTextSharp.text.Rectangle.NO_BORDER;
            document.Footer = footer;

            //Chapter对象和Section对象自动构建一个树：
            iTextSharp.text.Font f1 = new iTextSharp.text.Font();
            f1.SetStyle(iTextSharp.text.Font.BOLD);
            Paragraph cTitle = new Paragraph("This is chapter 1", f1);
            Chapter chapter = new Chapter(cTitle, 1);
            Paragraph sTitle = new Paragraph("This is section 1 in chapter 1", f1);
            Section section = chapter.AddSection(sTitle, 1);
            document.Add(chapter);



            ////构建了一个简单的表：

            //iTextSharp.text.Table aTable = new iTextSharp.text.Table(4, 4);

            //aTable.AutoFillEmptyCells = true;

            //aTable.AddCell("2.2", new Point(2, 2));

            //aTable.AddCell("3.3", new Point(3, 3));

            //aTable.AddCell("2.1", new Point(2, 1));

            //aTable.AddCell("1.3", new Point(1, 3));

            //document.Add(aTable);

            ////构建了一个不简单的表：

            //iTextSharp.text.Table table = new iTextSharp.text.Table(3);

            //table.BorderWidth = 1;

            //table.BorderColor = new iTextSharp.text.Color(0, 0, 255);

            //table.Cellpadding = 5;

            //table.Cellspacing = 5;

            //Cell cell = new Cell("header");

            //cell.Header = true;

            //cell.Colspan = 3;

            //table.AddCell(cell);

            //cell = new Cell("example cell with colspan 1 and rowspan 2");

            //cell.Rowspan = 2;

            //cell.BorderColor = new iTextSharp.text.Color(255, 0, 0);

            //table.AddCell(cell);

            //table.AddCell("1.1");

            //table.AddCell("2.1");

            //table.AddCell("1.2");

            //table.AddCell("2.2");

            //table.AddCell("cell test1");

            //cell = new Cell("big cell");

            //cell.Rowspan = 2;

            //cell.Colspan = 2;

            //cell.BackgroundColor = new iTextSharp.text.Color(0xC0, 0xC0, 0xC0);
            //table.AddCell(cell);
            //table.AddCell("cell test2");
            //// 改变了单元格“big cell”的对齐方式：
            //cell.HorizontalAlignment = Element.ALIGN_CENTER;
            //cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            //document.Add(table);



            //关闭Document

            document.Close();
        }
    }
}