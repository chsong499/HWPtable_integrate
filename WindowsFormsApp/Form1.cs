using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        // 파일 경로 변수 초기화
        string file_path = null;

        // 한 개의 uml 파일의 클래스(들)의 멤버(들) list
        List<List<string>> list_collection = new List<List<string>>();

        // 한 개의 uml 파일의 클래스(들) 이름 list
        List<string> title_collection = new List<string>(); 
        
        public Form1()
        {
            // 윈도우 최대화
            // this.WindowState = FormWindowState.Maximized; 

            // 레지스트리 주소에 hwp dll 등록, 접근 제한 메시지 나오지 않게 한다.
            InitializeComponent();
            {      
                const string HNCRoot = @"HKEY_Current_User\Software\HNC\HwpCtrl\Modules";
           
                //string myProjectPath = Path.GetFullPath(".\\");
                
                try
                {
                    if (Microsoft.Win32.Registry.GetValue(HNCRoot, "FilePathChecker", "Not Exist").Equals("Not Exist"))
                         Microsoft.Win32.Registry.SetValue(HNCRoot, "FilePathChecker", Environment.CurrentDirectory + "\\" + "FilePathCheckerModuleExample.dll");
                }
                catch { Microsoft.Win32.Registry.SetValue(HNCRoot, "FilePathChecker", Environment.CurrentDirectory + "\\" + "FilePathCheckerModuleExample.dll"); }
                bool result = axHwpCtrl1.RegisterModule("FilePathCheckDLL", "FilePathChecker");
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            title_collection.Clear(); // List 초기화
            list_collection.Clear();  // List 초기화

            XmlDocument xml = new XmlDocument();
            
            file_path = openFileDialog1.FileName;

            // 잘못된 파일 예외처리
            if (!file_path.Contains(".uml")) return;

            xml.Load(file_path);

            // xml 파일에 사용된 namespace 등록
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("XPD", "http://www.w3.org/1999/XSL/Transform");
            
            XmlNodeList xmlclass = xml.SelectNodes("//*[@type='UMLClass']", nsmgr);

            foreach (XmlNode xml2 in xmlclass)
            {
                // 형식(반환값), 이름을 저장하기 위한 list
                List<string> collection = new List<string>();

                // 클래스 이름 string을 title_collection에 따로 저장한다
                Console.WriteLine(xml2.FirstChild.InnerText);
                title_collection.Add(xml2.FirstChild.InnerText);

                // 멤버변수 저장
                XmlNodeList element2 = xml2.SelectNodes(".//*[@type='UMLAttribute']", nsmgr);
                foreach (XmlNode xnl in element2)
                {
                    // 형식(반환값) 저장
                    XmlNode type = null;
                    type = xnl.SelectSingleNode(".//*[@name='TypeExpression']", nsmgr);
                    if (type == null) collection.Add("void");  // 타입 표기를 안했을경우 void를 입력한다.
                    else
                    {
                        Console.WriteLine(type.FirstChild.InnerText + " type");
                        collection.Add(type.FirstChild.InnerText);
                    }
                    // 이름 저장
                    Console.WriteLine(xnl.FirstChild.InnerText);
                    collection.Add(xnl.FirstChild.InnerText);
                    
                } 

                // 멤버함수 저장
                XmlNodeList element = xml2.SelectNodes(".//*[@type='UMLOperation']", nsmgr);
                foreach (XmlNode xnl in element)
                {
                    // 형식(반환값) 저장
                    XmlNode type = null;
                    type = xnl.SelectSingleNode(".//*[@name='TypeExpression']", nsmgr);
                    if (type == null) collection.Add("void");  // 타입 표기를 안했을경우 void를 입력한다.
                    else
                    { 
                        Console.WriteLine(type.FirstChild.InnerText + " type");
                        collection.Add(type.FirstChild.InnerText);
                    }
                    // 이름 저장
                    Console.WriteLine(xnl.FirstChild.InnerText+"()");
                    collection.Add(xnl.FirstChild.InnerText + "()");
                    
                }
                // list_collection에 클래스별 멤버list를 저장
                list_collection.Add(collection);
            }
        }

        // 파일 가져오기 버튼
        private void button1_Click(object sender, EventArgs e)
        {
            file_path = null;
            string file_name = null;
            openFileDialog1.InitialDirectory = "C:\\"; 
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                file_path = openFileDialog1.FileName;
                file_name = file_path.Split('\\')[file_path.Split('\\').Length - 1];
                textBox1.Text = file_name;
            }
        }

        // 가운데 텍스트박스
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            file_path = openFileDialog1.FileName;
            axHwpCtrl1.Enter += new System.EventHandler(this.axHwpCtrl1_Enter);
        }

        // hwp COM
        private void axHwpCtrl1_Enter(object sender, EventArgs e)
        {
            // 파일을 불러오면 기존 hwp 지우기.
            axHwpCtrl1.Clear(1);

            Console.WriteLine(" h w p t e s t  . . .");

            // uml 파일의 클래스 갯수만큼 표 생성 script
            for (int i = 0; i < title_collection.Count; i++)
            {
                // 클래스명 표기
                var vStr0 = title_collection[i];
                HWPCONTROLLib.DHwpAction vAct0 = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("InsertText");
                HWPCONTROLLib.DHwpParameterSet vSet0 = (HWPCONTROLLib.DHwpParameterSet)vAct0.CreateSet();
                vAct0.GetDefault(vSet0);
                vSet0.SetItem("Text", vStr0);
                vAct0.Execute(vSet0);

                // 표 1개 생성
                int rowsCnt = ((list_collection[i].Count)/2+1); // row = (형식+이름)/2 + 헤더
                const int colsCnt = 4; // column

                HWPCONTROLLib.DHwpAction act = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("TableCreate");
                HWPCONTROLLib.DHwpParameterSet pset = (HWPCONTROLLib.DHwpParameterSet)act.CreateSet();
                act.GetDefault(pset);

                pset.SetItem("Rows", rowsCnt);
                pset.SetItem("Cols", colsCnt);

                HWPCONTROLLib.DHwpParameterArray rowsHeight = (HWPCONTROLLib.DHwpParameterArray)pset.CreateItemArray("RowHeight", rowsCnt);
                HWPCONTROLLib.DHwpParameterArray colsWidth = (HWPCONTROLLib.DHwpParameterArray)pset.CreateItemArray("ColWidth", colsCnt);

                for (int row_i = 0; row_i < rowsCnt; ++row_i) rowsHeight.SetItem(row_i, 2300);
                { for (int col_i = 0; col_i < colsCnt; ++col_i) colsWidth.SetItem(col_i, 7000); }

                act.Execute(pset);

                // 가운데 정렬 함수
                void allign_center() { 
                axHwpCtrl1.Run("ParagraphShapeAlignCenter");
                axHwpCtrl1.Run("TableVAlignCenter");
                axHwpCtrl1.Run("TableCellAlignCenterCenter");
                }
                allign_center();

                // 표 헤더 내용 작성 script
                var vStr = "구분";
                HWPCONTROLLib.DHwpAction vAct = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("InsertText");
                HWPCONTROLLib.DHwpParameterSet vSet = (HWPCONTROLLib.DHwpParameterSet)vAct.CreateSet();
                vAct.GetDefault(vSet);
                vSet.SetItem("Text", vStr);
                vAct.Execute(vSet);
                allign_center();

                axHwpCtrl1.Run("MoveRight");

                var vStr2 = "형식(반환 값)";
                HWPCONTROLLib.DHwpAction vAct2 = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("InsertText");
                HWPCONTROLLib.DHwpParameterSet vSet2 = (HWPCONTROLLib.DHwpParameterSet)vAct2.CreateSet();
                vAct2.GetDefault(vSet2);
                vSet2.SetItem("Text", vStr2);
                vAct2.Execute(vSet2);
                allign_center();

                axHwpCtrl1.Run("MoveRight");

                var vStr3 = "이름";
                HWPCONTROLLib.DHwpAction vAct3 = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("InsertText");
                HWPCONTROLLib.DHwpParameterSet vSet3 = (HWPCONTROLLib.DHwpParameterSet)vAct3.CreateSet();
                vAct3.GetDefault(vSet3);
                vSet3.SetItem("Text", vStr3);
                vAct3.Execute(vSet3);
                allign_center();

                axHwpCtrl1.Run("MoveRight");

                var vStr4 = "기능";
                HWPCONTROLLib.DHwpAction vAct4 = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("InsertText");
                HWPCONTROLLib.DHwpParameterSet vSet4 = (HWPCONTROLLib.DHwpParameterSet)vAct4.CreateSet();
                vAct4.GetDefault(vSet4);
                vSet4.SetItem("Text", vStr4);
                vAct4.Execute(vSet4);
                allign_center();

                // 아래 줄 첫째 칸으로
                axHwpCtrl1.Run("MoveDown");
                for(int i2=0;i2<3;i2++) axHwpCtrl1.Run("MoveLeft");

                // 함수, 변수의 첫번째 항목을 표기하는 flags (셀 병합 때 갯수 세는 기능도 한다)
                int first_flag = 0;
                int second_flag = 0;
                int third_flag = 0;

                //표 멤버변수, 멤버함수 내용 작성
                for (int j = 0; j < list_collection[i].Count;)
                {
                    Console.WriteLine(list_collection[i][j] + " 값을 표에 입력");

                    string member_kind;
                    
                    // 변수, 함수 구분
                    if (list_collection[i][j + 1].Contains("()"))
                    { member_kind = "함수"; third_flag=first_flag++; Console.WriteLine("함수 플래그"); }
                    else
                    { member_kind = "변수"; third_flag=second_flag++; Console.WriteLine("변수 플래그"); }
                    
                    // 구분 column
                    if (first_flag == 1 || second_flag == 1) // 항목이 존재할 경우 진입
                    {
                        if (third_flag == 0)
                        { 
                        var vStr5 = member_kind;
                        HWPCONTROLLib.DHwpAction vAct5 = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("InsertText");
                        HWPCONTROLLib.DHwpParameterSet vSet5 = (HWPCONTROLLib.DHwpParameterSet)vAct5.CreateSet();
                        vAct5.GetDefault(vSet5);
                        vSet5.SetItem("Text", vStr5);
                        vAct5.Execute(vSet5);
                        allign_center();

                            third_flag++; // 변수/함수 count가 1을 유지할 때 다른 변수나 함수가 계속 표기되지 않게 만든다.
                        }
                    }
                    axHwpCtrl1.Run("MoveRight");

                    // 형식(반환 값) column
                    var vStr6 = list_collection[i][j++];
                    HWPCONTROLLib.DHwpAction vAct6 = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("InsertText");
                    HWPCONTROLLib.DHwpParameterSet vSet6 = (HWPCONTROLLib.DHwpParameterSet)vAct6.CreateSet();
                    vAct6.GetDefault(vSet6);
                    vSet6.SetItem("Text", vStr6);
                    vAct6.Execute(vSet6);
                    allign_center();
                    
                    axHwpCtrl1.Run("MoveRight");

                    // 이름 column
                    var vStr7 = list_collection[i][j++];
                    HWPCONTROLLib.DHwpAction vAct7 = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("InsertText");
                    HWPCONTROLLib.DHwpParameterSet vSet7 = (HWPCONTROLLib.DHwpParameterSet)vAct7.CreateSet();
                    vAct7.GetDefault(vSet7);
                    vSet7.SetItem("Text", vStr7);
                    vAct7.Execute(vSet7);
                    allign_center();

                    // 기능 column
                    axHwpCtrl1.Run("MoveRight");
                    allign_center();
                    
                    // 아래 줄 첫째 칸으로
                    axHwpCtrl1.Run("MoveDown");
                    if(j < list_collection[i].Count) 
                    for (int i2 = 0; i2 < 3; i2++) axHwpCtrl1.Run("MoveLeft");
                }

                // 멤버변수/멤버함수가 2개 이상이라 merge가 필요한 경우 script
                if (first_flag > 1 || second_flag > 1)
                {
                    // 표 가장 아래row 첫째 column으로 이동
                    axHwpCtrl1.Run("MoveUp");
                    axHwpCtrl1.Run("TableLeftCell");
                    axHwpCtrl1.Run("TableLeftCell");
                    axHwpCtrl1.Run("TableLeftCell");

                    // first_flag = 함수 갯수, second_flag = 변수 갯수
                    // 아래에서부터 같은 영역을 블록 설정해서 위로 올라간다. 그러니 함수부터 merge를 시작한다.
                    // 함수,변수가 2개 이상인 경우, 1개인 경우, 0개인 경우 각각 merge 방식이 다르다.

                    if (first_flag > 1)
                    {
                        axHwpCtrl1.Run("TableCellBlockExtendAbs");
                        for (; first_flag > 1; --first_flag)
                        { axHwpCtrl1.Run("TableUpperCell"); }
                        axHwpCtrl1.Run("TableMergeCell");
                        axHwpCtrl1.Run("Cancel");

                        axHwpCtrl1.Run("MoveUp");
                    }
                    else if (first_flag == 1)
                    {
                        axHwpCtrl1.Run("MoveUp");
                    }

                    if (second_flag > 1)
                    {
                        axHwpCtrl1.Run("TableCellBlockExtendAbs");
                        for (; second_flag > 1; --second_flag)
                        { axHwpCtrl1.Run("TableUpperCell"); }
                        axHwpCtrl1.Run("TableMergeCell");
                        axHwpCtrl1.Run("Cancel");
                    }
                    else if (second_flag == 1)
                    {
                    }
                }

                axHwpCtrl1.MovePos(0,0,0); // 커서 이동 함수. 없으면 표 생성 오류.
                
                // 보기 편하게 공백 삽입 script
                var vStr00 = "\r";
                HWPCONTROLLib.DHwpAction vAct00 = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("InsertText");
                HWPCONTROLLib.DHwpParameterSet vSet00 = (HWPCONTROLLib.DHwpParameterSet)vAct00.CreateSet();
                vAct00.GetDefault(vSet00);
                vSet00.SetItem("Text", vStr00);
                vAct00.Execute(vSet00);
            }
            
            // 다음 파일에 쓰기 위해서 clear()
            list_collection.Clear();
            title_collection.Clear();
            axHwpCtrl1.Enter-= new System.EventHandler(this.axHwpCtrl1_Enter);
            
        }

        // 저장하기 버튼
        private void button2_Click(object sender, EventArgs e)
        {
            // 파일명_class.hwp 생성
            string file_name = file_path.Split('\\')[file_path.Split('\\').Length - 1];
            string new_name = file_path.Split('.')[file_path.Split('.').Length-2] + "_class.hwp";
            axHwpCtrl1.SaveAs(new_name, "HWP", null);
        }

        private void axHwpCtrl1_NotifyMessage(object sender, AxHWPCONTROLLib._DHwpCtrlEvents_NotifyMessageEvent e)
        {
        }
    }
}
