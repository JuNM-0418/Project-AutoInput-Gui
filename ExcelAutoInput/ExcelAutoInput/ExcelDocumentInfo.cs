using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAutoInput
{
    internal class ExcelDocumentInfo
    {
        private Excel.Application application = null;
        private string FileName = null;
        private Excel.Workbook WorkBook = null;
        private List<Excel.Worksheet> WorkSheetList = null; 
        private List<Excel.Worksheet> SelectedSheetList = null;
        private List<Excel.Worksheet> SelectedSurveySheetList = null;
        private List<string> ImgFolderPathList = null;
        private int LocationRowTop = 3; // 페이지 내 첫번째 시작 행
        private int LocationColImg = 2; // 사진 삽입 고정 열
        private int ImgCycle = 1;   // 각 동의 폴더에서 넣을 사진의 순서
        private int SurveyDataCycle = 1; // 조사표 내용이 삽입되는 순서
        private int SurveyDataNumber = 6; // 조사표 시트에서 처음으로 데이터가 들어있는 행 넘버
        private int PageNum = 0;


        // 다음 행의 위치를 반환하는 함수
        public int NextLocationRow(int locationRow)
        {
            locationRow += 32;
            return locationRow;
        }

        // 사진 넣어주는 함수
        public void InputImg(Excel.Worksheet workSheet, string imgFolderPath)
        {
            this.LocationRowTop = 3;
            this.LocationColImg = 2;
            this.ImgCycle = 1;

            for (int i = 0; i < this.GetPageNum(); i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    Excel.Range range = workSheet.Cells[this.LocationRowTop + (j * 10), this.LocationColImg];
                    string imgPath = imgFolderPath + "\\" + (this.ImgCycle).ToString() + ".jpg";
                    workSheet.Shapes.AddPicture(@imgPath, 0, (Microsoft.Office.Core.MsoTriState)1, range.Left, range.Top, (float)246.68, (float)184.28);
                    this.ImgCycle++;
                }
                this.LocationRowTop = this.NextLocationRow(this.LocationRowTop);
            }
        }

        // 화살표를 넣어주는 함수
        public void InputArrow(Excel.Worksheet workSheet)
        {
            this.LocationRowTop = 3;

            for (int i = 0; i < this.GetPageNum(); i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    workSheet.Range["Z1:AF4"].Copy(workSheet.Range["U" + (this.LocationRowTop + (j * 10))]);
                    this.ImgCycle++;
                }

                this.LocationRowTop = this.NextLocationRow(this.LocationRowTop);
            }
        }

        // 설명 번호와 사진 번호가 맞는지 확인해주는 수식을 넣어주는 함수
        public void InputCheckImgNumFunction(Excel.Worksheet workSheet)
        {
            this.LocationRowTop = 3;
            for (int i = 0; i < this.GetPageNum(); i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    workSheet.Cells[this.LocationRowTop + 1 + (j * 10), 21].Value = "=IF(MID(O" + (this.LocationRowTop + 1 + (j * 10)) + ",SEARCH(\".\",O" + (this.LocationRowTop + 1 + (j * 10)) + ",1),3)=MID(W" + (this.LocationRowTop + (j * 10) + 8) + ",SEARCH(\".\",W" + (this.LocationRowTop + (j * 10) + 8) + ",1),3),\"\",\"번호확인\")";
                    workSheet.Cells[this.LocationRowTop + 1 + (j * 10), 21].Font.Color = -16776961;
                }
                this.LocationRowTop = this.NextLocationRow(this.LocationRowTop);
            }
        }

        // 조사표 내용을 연동시키는 수식을 넣어주는 함수
        public void InputSurveyData(Excel.Worksheet surveyWorkSheet, Excel.Worksheet imgWorkSheet)
        {
            string surveyFilter = "사진.";
            string surveyEnd = "끝";
            int surveyCycle = 6;
            this.LocationRowTop = 3;

            for (int i = 0; i < this.GetPageNum(); i++)
            {
                for (int j = 0; j <3;)
                {
                    string tmp = surveyWorkSheet.Cells[surveyCycle, 18].Text.ToString();

                    if(tmp.Length > 0)
                    {
                        if (tmp.IndexOf(surveyFilter) != -1)
                        {
    
                            imgWorkSheet.Cells[this.LocationRowTop + 1 + (j * 10), 17].Formula = ("='" + surveyWorkSheet.Name + "'!B" + surveyCycle.ToString() + "&CHAR(10)&'" + surveyWorkSheet.Name + "'!D" + surveyCycle.ToString());
                            if (surveyWorkSheet.Cells[surveyCycle, 16].Text.ToString() != "-")
                            {
                                imgWorkSheet.Cells[this.LocationRowTop + 4 + (j * 10), 15].Formula = ("='" + surveyWorkSheet.Name + "'!P" + surveyCycle.ToString() + "&'" + surveyWorkSheet.Name + "'!G" + surveyCycle.ToString());
                            }
                            else
                            {
                                imgWorkSheet.Cells[this.LocationRowTop + 4 + (j * 10), 15].Formula = ("='" + surveyWorkSheet.Name + "'!G" + surveyCycle.ToString());
                            }
                            imgWorkSheet.Cells[this.LocationRowTop + 8 + (j * 10), 15].Formula = ("='" + surveyWorkSheet.Name + "'!J" + surveyCycle.ToString());
                            imgWorkSheet.Cells[this.LocationRowTop + 8 + (j * 10), 17].Formula = ("='" + surveyWorkSheet.Name + "'!L" + surveyCycle.ToString());
                            imgWorkSheet.Cells[this.LocationRowTop + 8 + (j * 10), 19].Formula = ("='" + surveyWorkSheet.Name + "'!N" + surveyCycle.ToString());
                            j++;
                            surveyCycle++;
                        }
                        else if(tmp.IndexOf(surveyEnd) != -1)
                        {
                            break;
                        }
                        else
                        {
                            surveyCycle++;
                        }
                    }
                    else
                    {
                        surveyCycle++;
                    }

                    
                }
                this.LocationRowTop = this.NextLocationRow(this.LocationRowTop);
            }
            

        }

        // 이미지의 갯수가 3의 배수가 아니면 복사하는 함수
        public void DuplicateImg(string folderPath)
        {
            DirectoryInfo di = new DirectoryInfo(@folderPath);
            IEnumerable<FileInfo> imgList = di.EnumerateFiles("*.jpg", SearchOption.AllDirectories);
            if (imgList.Count() % 3 != 0)
            {
                int addImgNum = 3 - (imgList.Count() % 3);
                for (int i = 1; i <= addImgNum; i++)
                {
                    File.Copy(imgList.First().FullName, imgList.First().Directory + "\\" + (imgList.Count() + 1) + ".jpg");
                }
            }
            imgList = di.EnumerateFiles("*.jpg", SearchOption.AllDirectories);
        }
   
        public void SetFileName(string fileName)
        {
            this.FileName = fileName;
        }
        public string GetFileName()
        {
            return this.FileName;
        }


        public Excel.Workbook OpenWorkbook()
        {
            this.WorkBook = application.Workbooks.Open(GetFileName());
            return this.WorkBook;
        }


        public void SetWorkSheetList(Excel.Workbook workBook)
        {
            this.WorkSheetList = new List<Excel.Worksheet>();
            for (int i = 1; i <= this.WorkBook.Sheets.Count; i++)
            {
                this.WorkSheetList.Add(workBook.Sheets.get_Item(i));
            }
        }
        public List<Excel.Worksheet> GetWorkSheetList()
        {
            return this.WorkSheetList;
        }


        public void SetSelectedSheetList(Excel.Worksheet selectedSheet)
        {
            if (SelectedSheetList == null)
            {
                SelectedSheetList = new List<Excel.Worksheet>();
            }
            this.SelectedSheetList.Add(selectedSheet);

        }
        public List<Excel.Worksheet> GetSelectedSheetList()
        {
            return this.SelectedSheetList;
        }

        // 조사표 시트 Set, Get 함수
        public void SetSelectedSurveySheetList(Excel.Worksheet selectedSurveySheet)
        {
            if (SelectedSurveySheetList == null)
            {
                SelectedSurveySheetList = new List<Excel.Worksheet>();
            }
            this.SelectedSurveySheetList.Add(selectedSurveySheet);

        }
        public List<Excel.Worksheet> GetSelectedSurveySheetList()
        {
            return this.SelectedSurveySheetList;
        }


        public void SetSelectedImgFolderPathList(string folderPath)
        {
            if (ImgFolderPathList == null)
            {
                ImgFolderPathList = new List<string>();
            }
            this.ImgFolderPathList.Add(folderPath);

        }
        public List<string> GetSelectedImgFolderPathList()
        {
            return this.ImgFolderPathList;
        }

        public void SetPageNum(string imgFolderPath)
        {
            DirectoryInfo di = new DirectoryInfo(imgFolderPath);
            this.PageNum = di.EnumerateFiles("*.jpg", SearchOption.AllDirectories).Count() / 3;
        }
        public int GetPageNum()
        {
            return this.PageNum;
        }

        public void ShowExcel()
        {
            this.application.Visible = true;
        }

        public ExcelDocumentInfo()
        {
            this.application = new Excel.Application();
        }
        ~ExcelDocumentInfo()
        {
            this.application.Visible = true;
        }

    }
}
