using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;



namespace ExcelAutoInput
{
    public partial class Form1 : Form
    {

        private ExcelDocumentInfo excelFileInfo = null;
        private DirectoryInfo di = null;


        public Form1()
        {
            InitializeComponent();
            btnSelectSheet.Enabled = false;
            btnSelectRootFolder.Enabled = false;
            btnSelectImgFolder.Enabled = false;
            btnInputImage.Enabled = false;
            btnSelectSurveySheet.Enabled = false;
            btnSelectSurveySheetDone.Enabled = false;
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            excelFileInfo = new ExcelDocumentInfo();
            openFileDialog = new OpenFileDialog();
            Console.Write(openFileDialog.FileName);
            openFileDialog.ShowDialog();
            excelFileInfo.SetFileName(openFileDialog.FileName);

            if(excelFileInfo.GetFileName() != "")
            {
                excelFileInfo.SetWorkSheetList(excelFileInfo.OpenWorkbook());
                for (int i = 0; i < excelFileInfo.GetWorkSheetList().Count; i++)
                {
                    sheetListBox.Items.Add(excelFileInfo.GetWorkSheetList()[i].Name);
                }
                btnOpenFile.Enabled = false;
                btnSelectSheet.Enabled = true;
                
            }
            else
            {
                MessageBox.Show("엑셀 파일을 선택해주세요", "오류", MessageBoxButtons.OK ,MessageBoxIcon.Error);
            }
            
        }
        private void btnSelectSheet_Click(object sender, EventArgs e)
        {
            foreach (string checkedSheet in sheetListBox.CheckedItems)
            {
                foreach (Excel.Worksheet workSheet in excelFileInfo.GetWorkSheetList())
                    if (workSheet.Name == checkedSheet)
                    {
                        excelFileInfo.SetSelectedSheetList(workSheet);
                    }
            }
            if(excelFileInfo.GetSelectedSheetList() != null)
            {
                sheetListBox.Enabled = false;
                btnSelectSheet.Enabled = false;
                btnSelectRootFolder.Enabled = true;
            }
            else
            {
                MessageBox.Show("엑셀 시트를 선택해주세요", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }


        private void btnSelectRootFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
            if(folderBrowserDialog.SelectedPath != "")
            {
                di = new DirectoryInfo(@folderBrowserDialog.SelectedPath);
                foreach (DirectoryInfo directory in di.EnumerateDirectories())
                {
                    imgPathListBox.Items.Add(directory.Name);
                }

                btnSelectRootFolder.Enabled = false;
                btnSelectImgFolder.Enabled = true;
                
            }
            else
            {
                MessageBox.Show("작업 폴더를 선택해주세요", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void btnInputImage_Click(object sender, EventArgs e)
        {
            btnInputImage.Enabled = false;
            btnExit.Enabled = false;
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Minimum = 0;
            progressBar.Maximum = imgPathListBox.CheckedItems.Count;
            progressBar.Value = 0;
            progressBar.Step = 1;

            for (int i = 0; i < imgPathListBox.CheckedItems.Count; i++)
            {
                progressBar.PerformStep();

                // 이미지의 갯수가 3의 배수가 아니면 복사하는 함수
                excelFileInfo.DuplicateImg(excelFileInfo.GetSelectedImgFolderPathList()[i]);
                // 페이지 갯수 설정 함수
                excelFileInfo.SetPageNum(excelFileInfo.GetSelectedImgFolderPathList()[i]);
                // 사진 넣어주는 함수
                excelFileInfo.InputImg(excelFileInfo.GetSelectedSheetList()[i], excelFileInfo.GetSelectedImgFolderPathList()[i]);
                // 화살표를 넣어주는 함수
                excelFileInfo.InputArrow(excelFileInfo.GetSelectedSheetList()[i]);
                // 설명 번호와 사진 번호가 맞는지 확인하는 수식을 넣어주는 함수
                excelFileInfo.InputCheckImgNumFunction(excelFileInfo.GetSelectedSheetList()[i]);
                // 조사표 내용을 연동시켜주는 함수
                excelFileInfo.InputSurveyData(excelFileInfo.GetSelectedSurveySheetList()[i], excelFileInfo.GetSelectedSheetList()[i]);

            }
            MessageBox.Show("완료되었습니다.");
            btnExit.Enabled = true;
        }


        private void btnExit_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnSelectImgFolder_Click(object sender, EventArgs e)
        {
            foreach (string folderPath in imgPathListBox.CheckedItems)
            {
                excelFileInfo.SetSelectedImgFolderPathList(di.FullName + "\\" + folderPath);
            }
            imgPathListBox.Enabled = false;
            btnSelectImgFolder.Enabled = false;
            btnSelectSurveySheet.Enabled = true;
            
        }

        private void btnSelectSurveySheet_Click(object sender, EventArgs e)
        {
            btnSelectSurveySheet.Enabled = false;
            btnSelectSurveySheetDone.Enabled = true;

            excelFileInfo.SetWorkSheetList(excelFileInfo.OpenWorkbook());
            for (int i = 0; i < excelFileInfo.GetWorkSheetList().Count; i++)
            {
                surveySheetListBox.Items.Add(excelFileInfo.GetWorkSheetList()[i].Name);
            }
        }

        private void btnSelectSurveySheetDone_Click(object sender, EventArgs e)
        {
            btnSelectSurveySheetDone.Enabled = false;


            foreach (string checkedSheet in surveySheetListBox.CheckedItems)
            {
                foreach (Excel.Worksheet workSheet in excelFileInfo.GetWorkSheetList())
                    if (workSheet.Name == checkedSheet)
                    {
                        excelFileInfo.SetSelectedSurveySheetList(workSheet);
                    }
            }
            if (excelFileInfo.GetSelectedSurveySheetList() != null)
            {
                surveySheetListBox.Enabled = false;
                btnSelectSurveySheet.Enabled = false;
                btnSelectSurveySheetDone.Enabled = false;
                btnInputImage.Enabled = true;
            }
            else
            {
                MessageBox.Show("엑셀 시트를 선택해주세요", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
    }
}
