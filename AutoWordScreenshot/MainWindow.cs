using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoWordScreenshot {
    public partial class MainWindow : Form {

        private Utils utils = new Utils();
        private string selectedDocumentPath = string.Empty;

        private Label selectedDocumentLabel;
        private Button selectButton;
        private Button captureButton;

        public MainWindow() {
            InitializeComponent();
            this.ClientSize = new Size(320, 80);

            selectedDocumentLabel = new Label {
                Location = new System.Drawing.Point(170, 17),
                Text = ""
            };
            selectButton = new Button {
                Size = new Size(150, 25),
                Location = new System.Drawing.Point(12, 10),
                Text = "Select Document"
            };
            selectButton.Click += new EventHandler(SelectButton_Click);
            captureButton = new Button {
                Size = new Size(296, 25),
                Location = new System.Drawing.Point(12, 40),
                Text = "Capture Screenshot"
            };
            captureButton.Click += new EventHandler(CaptureButton_Click);

            Controls.Add(selectedDocumentLabel);
            Controls.Add(captureButton);
            Controls.Add(selectButton);
        }

        //! Click Events

        private void SelectButton_Click(object sender, EventArgs e) {
            using (OpenFileDialog openFileDialog = new OpenFileDialog()) {
                openFileDialog.Filter = "Word Documents|*.docx";

                if (openFileDialog.ShowDialog() != DialogResult.OK) {
                    return;
                }

                if (String.IsNullOrEmpty(openFileDialog.FileName) || !openFileDialog.CheckFileExists || !openFileDialog.FileName.EndsWith(".docx")) {
                    MessageBox.Show("Please select a valid word document.");
                    return;
                }

                selectedDocumentPath = openFileDialog.FileName;
                selectedDocumentLabel.Text = openFileDialog.SafeFileName;
            }
        }

        private async void CaptureButton_Click(object sender, EventArgs e) {
            if (string.IsNullOrEmpty(selectedDocumentPath)) {
                MessageBox.Show("Please select a valid word document.");
                return;
            }

            this.Opacity = 0;

            using (Bitmap bitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)) {
                using (Graphics graphics = Graphics.FromImage(bitmap)) {
                    graphics.CopyFromScreen(0, 0, 0, 0, bitmap.Size, CopyPixelOperation.SourceCopy);
                    this.Opacity = 1;

                    string tempCopyFile = utils.SaveTempImage(bitmap);

                    try {
                        await Task.Run(() => {
                            using (WordprocessingDocument document = utils.GetDocument(selectedDocumentPath)) {
                                if (document == null) {
                                    return;
                                }

                                ImagePart imagePart = document.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);

                                using (FileStream stream = new FileStream(tempCopyFile, FileMode.Open)) {
                                    imagePart.FeedData(stream);
                                }

                                utils.SaveImageToDocument(document, document.MainDocumentPart.GetIdOfPart(imagePart), bitmap.Width, bitmap.Height);
                            }
                        });
                    } catch (Exception ex) {
                        Console.WriteLine(ex.Message);
                    } finally {
                        File.Delete(tempCopyFile);
                    }  
                }
            }
        }
    }
}
