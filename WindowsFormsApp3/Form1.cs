using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;

namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        private static Mutex mutex = new Mutex(); // mutex for 재귀함수
        int recur_exception = 0; // 재귀함수 반복 카운트 초기화
        string folder_path = null; // 디렉토리 경로 초기화

        class Element // 파일 정보 클래스
        {
            public string filename="";
            public string version = "1.0"; // 버전은 1.0으로 통일한다.
            public string size = "";
            public string checksum = "";
            public string birthday = "";
            public string lines = "";
            public Boolean flag = false;
            public Boolean flagcheck() { return flag; } // 인스턴스에 파일 정보 대신 저장위치를 입력할 때는 flag에 true를 입력한다.
        }

        List<Element> ele_collection = new List<Element>(); // List for Element Class
        
        public Form1()
        {
            // 윈도우 최대화
            // this.WindowState = FormWindowState.Maximized; 

            // 레지스트리 주소에 hwp dll 등록, 접근 제한 메시지 나오지 않게 만든다.
            InitializeComponent();
            {      
                const string HNCRoot = @"HKEY_Current_User\Software\HNC\HwpCtrl\Modules";
                
                try
                {
                    if (Microsoft.Win32.Registry.GetValue(HNCRoot, "FilePathChecker", "Not Exist").Equals("Not Exist"))
                         Microsoft.Win32.Registry.SetValue(HNCRoot, "FilePathChecker", Environment.CurrentDirectory + "\\" + "FilePathCheckerModuleExample.dll");
                }
                catch { Microsoft.Win32.Registry.SetValue(HNCRoot, "FilePathChecker", Environment.CurrentDirectory + "\\" + "FilePathCheckerModuleExample.dll"); }
                bool result = axHwpCtrl1.RegisterModule("FilePathCheckDLL", "FilePathChecker");
            }
        }

        // 가져오기 버튼
        private void button1_Click(object sender, EventArgs e)
        {
            folder_path = null;
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            folder_path = dialog.SelectedPath;
            
            if (folder_path.Length > 0)
            {
                textBox1.Text = folder_path;
            }
        }

        // 가운데 텍스트박스
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Console.Write("First");
            //folderBrowserDialog1_work(sender, e);
            //axHwpCtrl1.Enter += new System.EventHandler(this.axHwpCtrl1_Enter);
        }

        // hwp COM
        private void axHwpCtrl1_Enter(object sender, EventArgs e)
        {
            // 폴더를 가져오면 기존 hwp 지우기.
            axHwpCtrl1.Clear(1);

            int rowsCnt = (ele_collection.Count + 1); // row = 항목+헤더
            const int colsCnt = 8; // column

            HWPCONTROLLib.DHwpAction act = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("TableCreate");
            HWPCONTROLLib.DHwpParameterSet pset = (HWPCONTROLLib.DHwpParameterSet)act.CreateSet();
            act.GetDefault(pset);

            pset.SetItem("Rows", rowsCnt);
            pset.SetItem("Cols", colsCnt);

            HWPCONTROLLib.DHwpParameterArray rowsHeight = (HWPCONTROLLib.DHwpParameterArray)pset.CreateItemArray("RowHeight", rowsCnt);
            HWPCONTROLLib.DHwpParameterArray colsWidth = (HWPCONTROLLib.DHwpParameterArray)pset.CreateItemArray("ColWidth", colsCnt);

            for (int row_i = 0; row_i < rowsCnt; ++row_i) rowsHeight.SetItem(row_i, 2300);
            for (int col_i = 0; col_i < colsCnt; ++col_i) colsWidth.SetItem(col_i, 4200);

            act.Execute(pset);

            // 한글 표 string 입력 script 함수
            void write_string(string s)
            {
                axHwpCtrl1.Run("ParagraphShapeAlignLeft"); //왼쪽 정렬
                var vStr00 = s;
                HWPCONTROLLib.DHwpAction vAct00 = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("InsertText");
                HWPCONTROLLib.DHwpParameterSet vSet00 = (HWPCONTROLLib.DHwpParameterSet)vAct00.CreateSet();
                vAct00.GetDefault(vSet00);
                vSet00.SetItem("Text", vStr00);
                vAct00.Execute(vSet00);
            }

            // 헤더 작성
            string[] headers = { "순번", "파일명", "버전", "크기 (Byte)", "첵섬", "생성일자", "라인수" /*, "기능 설명" */ };
            foreach(string i2 in headers)
            {
                write_string(i2);
                axHwpCtrl1.Run("MoveRight");
            }
            
            write_string("기능 설명");

            // 아래 줄 첫째 칸으로
            axHwpCtrl1.Run("MoveDown");
            for(int i2=0;i2<7;i2++) axHwpCtrl1.Run("MoveLeft");
            
            int i = 1;

            foreach (Element ele in ele_collection)
            {

                // 폴더일 경우
                if (ele.flagcheck())
                {
                    i = 1; //순번 작성

                    // mutex
                    mutex.WaitOne();

                    Console.WriteLine(ele.filename);
                    write_string("저장위치: " + ele.filename);

                    axHwpCtrl1.Run("TableCellBlockExtendAbs");
                    for (int j = 0; j < 7; j++) { axHwpCtrl1.Run("TableRightCell"); };
                    axHwpCtrl1.Run("TableMergeCell");
                    axHwpCtrl1.Run("Cancel");

                    // 저장위치가 한 줄을 넘어가면 MoveDown은 같은 셀 다음 줄로 이동하므로 TableLowerCell을 해야 한다.
                    axHwpCtrl1.Run("TableLowerCell"); 
                    mutex.ReleaseMutex();
                }
                // 항목 작성
                else
                {
                    write_string("" + (i++));
                    axHwpCtrl1.Run("MoveRight");

                    write_string(ele.filename);
                    axHwpCtrl1.Run("MoveRight");

                    write_string(ele.version);
                    axHwpCtrl1.Run("MoveRight");

                    write_string(ele.size);
                    axHwpCtrl1.Run("MoveRight");

                    write_string(ele.checksum);
                    axHwpCtrl1.Run("MoveRight");

                    write_string(ele.birthday);
                    axHwpCtrl1.Run("MoveRight");

                    write_string(ele.lines);

                    // 아랫줄 첫째칸으로 이동
                    axHwpCtrl1.Run("MoveDown");
                    for (int k = 0; k < 6; k++) axHwpCtrl1.Run("MoveLeft");
                }
            }

            axHwpCtrl1.MovePos(0, 0, 0); // 커서 이동 함수. 없으면 표 생성 오류.

            // 다음 파일에 쓰기 위해서 clear()
            ele_collection.Clear();
            //axHwpCtrl1.Enter -= new System.EventHandler(this.axHwpCtrl1_Enter);
        }

        // 저장하기 버튼
        private void button2_Click(object sender, EventArgs e)
        {
            // 디렉토리명_directory.hwp 생성
            string[] new_name_array = folder_path.Split('\\');
            string upper_path = "";
            int i = 0;
            for (; i < new_name_array.Length-1;i++)
                { upper_path += new_name_array[i]+'\\'; }
            axHwpCtrl1.SaveAs(upper_path + new_name_array[i] + "_directory.hwp", "HWP", null);
        }

        private void folderBrowserDialog1_work(object sender, EventArgs e)
        {
            ele_collection.Clear(); // bugfix

            DirectoryInfo path = new DirectoryInfo(folder_path);

            int table_flag = 0; // 저장위치 반복 저장을 막기 위한 flag

            // 재귀함수 실행
            GetFile(path);

            // 재귀함수 카운트 초기화
            recur_exception = 0; 

            // 하위 디렉토리를 탐색하는 재귀함수 정의
            void GetFile(DirectoryInfo di)
            {
                // 하위 폴더 목록들의 정보를 가져온다.
                DirectoryInfo[] di_sub = di.GetDirectories(); 

                if (di_sub.Length > 0)
                {
                    recur_exception++;

                    // 하위 폴더목록 스캔
                    foreach (DirectoryInfo di1 in di_sub) 
                    {
                        // mutex
                        mutex.WaitOne();
                        if (recur_exception > 30 && checkBox1.Checked) return; // 재귀 제한
                        
                        GetFile(di1);
                        mutex.ReleaseMutex();
                    }
                }

                // 선택 폴더의 파일 목록 스캔. 파일 없으면 해당 안됨(파일이 없는 저장위치의 입력은 걱정할 필요 없다)
                foreach (FileInfo File in di.GetFiles())
                {
                    if (table_flag == 0) // 저장위치 instance는 list에 한번만 입력
                    {
                        // instance의 flag에 true 지정하고 ele_collection에 추가
                        Element address = new Element();
                        address.flag = true;
                        address.filename = File.DirectoryName+"\\"; // 저장위치 문자열 입력
                        Console.WriteLine(address.filename+" 추가");
                        ele_collection.Add(address);
                        table_flag = 1; // table_flag에 1 지정
                    }

                    // 파일 확장자 체크 변수
                    string file_extension = File.Name.Split('\\')[File.Name.Split('\\').Length - 1];
                    file_extension = file_extension.Split('.')[file_extension.Split('.').Length-1];

                    // 확장자 == c,cpp,h,hpp
                    if ((file_extension == "c") || (file_extension == "h") || (file_extension == "cpp") || (file_extension == "hpp"))
                    {
                        // 새 instance에 파일정보 입력, ele_collection에 추가
                        Console.WriteLine(File.Name + "추가");

                        Element file_element = new Element();

                        file_element.filename = File.Name;

                        file_element.size = string.Format("{0:n0}", File.Length);

                        // checksum 구하기
                        //byte[] fileBytes 생성
                        byte[] fileBytes = null;
                        FileStream fileStream = new FileStream(File.FullName, FileMode.Open, FileAccess.Read);
                        fileBytes = new byte[fileStream.Length];
                        fileStream.Read(fileBytes, 0, fileBytes.Length);
                        // fileBytes 이용해서 checksum 계산
                        uint crc32 = 0xFFFFFFFF;
                        for (uint nLoop = 0; nLoop < File.Length; nLoop++)
                        { crc32 = (crc32 >> 8) ^ (defaultTable[(byte)(((uint)fileBytes[nLoop]) ^ (crc32 & 0x000000FF))]); }
                        crc32 = ~crc32; // 뒤집는다
                        file_element.checksum = "" + crc32.ToString("X2"); // 16진수 hexadecimal로 변환해서 출력

                        // 생성 일자 구하기
                        string birthday_temp = "" + File.CreationTime;
                        birthday_temp = birthday_temp.Remove(10);
                        birthday_temp = birthday_temp.Replace("-", ".");
                        List<int> zero_list = new List<int>();
                        file_element.birthday = birthday_temp;

                        // 라인 수 구하기
                        var lineCount = System.IO.File.ReadLines(@File.FullName).Count();
                        file_element.lines = "" + lineCount;

                        //ele_collection에 instance 추가
                        ele_collection.Add(file_element);
                    }

                    // 확장자 == 이미지 포맷
                    else if ((file_extension == "jpg") || (file_extension == "JPG") || (file_extension == "jpeg")
                        || (file_extension == "JPEG") || (file_extension == "bmp") || (file_extension == "BMP"))
                    {
                        // 새 instance에 파일정보 입력, ele_collection에 추가
                        Console.WriteLine(File.Name + "추가");

                        Element file_element = new Element();

                        file_element.filename = File.Name;

                        file_element.size = string.Format("{0:n0}", File.Length);

                        // checksum 구하기
                        //byte[] fileBytes 생성
                        byte[] fileBytes = null;
                        FileStream fileStream = new FileStream(File.FullName, FileMode.Open, FileAccess.Read);
                        fileBytes = new byte[fileStream.Length];
                        fileStream.Read(fileBytes, 0, fileBytes.Length);
                        // fileBytes 이용해서 checksum 계산
                        uint crc32 = 0xFFFFFFFF;
                        for (uint nLoop = 0; nLoop < File.Length; nLoop++)
                        { crc32 = (crc32 >> 8) ^ (defaultTable[(byte)(((uint)fileBytes[nLoop]) ^ (crc32 & 0x000000FF))]); }
                        crc32 = ~crc32; // 뒤집는다
                        file_element.checksum = "" + crc32.ToString("X2"); // 16진수 hexadecimal로 변환해서 출력

                        // 생성 일자 구하기
                        string birthday_temp = "" + File.CreationTime;
                        birthday_temp = birthday_temp.Remove(10);
                        birthday_temp = birthday_temp.Replace("-", ".");
                        List<int> zero_list = new List<int>();
                        file_element.birthday = birthday_temp;

                        // 사진 크기, 비트 수준
                        var img = Image.FromFile(File.FullName);
                        string image_size = img.Width + "×" + img.Height;
                        string bit_temp = "" + img.PixelFormat;
                        bit_temp = bit_temp.Substring(6, 2);
                        file_element.lines = "" + image_size + '\r' + bit_temp + "bit";

                        //ele_collection에 instance 추가
                        ele_collection.Add(file_element);
                    }

                    // 확장자 == wav
                    else if ((file_extension == "wav"))
                    {
                        // 새 instance에 파일정보 입력, ele_collection에 추가
                        Console.WriteLine(File.Name + "추가");

                        Element file_element = new Element();

                        file_element.filename = File.Name;

                        file_element.size = string.Format("{0:n0}", File.Length);

                        // checksum 구하기
                        //byte[] fileBytes 생성
                        byte[] fileBytes = null;
                        FileStream fileStream = new FileStream(File.FullName, FileMode.Open, FileAccess.Read);
                        fileBytes = new byte[fileStream.Length];
                        fileStream.Read(fileBytes, 0, fileBytes.Length);
                        // fileBytes 이용해서 checksum 계산
                        uint crc32 = 0xFFFFFFFF;
                        for (uint nLoop = 0; nLoop < File.Length; nLoop++)
                        { crc32 = (crc32 >> 8) ^ (defaultTable[(byte)(((uint)fileBytes[nLoop]) ^ (crc32 & 0x000000FF))]); }
                        crc32 = ~crc32; // 뒤집는다
                        file_element.checksum = "" + crc32.ToString("X2"); // 16진수 hexadecimal로 변환해서 출력

                        // 생성 일자 구하기
                        string birthday_temp = "" + File.CreationTime;
                        birthday_temp = birthday_temp.Remove(10);
                        birthday_temp = birthday_temp.Replace("-", ".");
                        List<int> zero_list = new List<int>();
                        file_element.birthday = birthday_temp;

                        // 비트레이트
                        int bitrate;
                        var f = File.OpenRead();
                        f.Seek(28, SeekOrigin.Begin);
                        byte[] val = new byte[4];
                        f.Read(val, 0, 4);
                        bitrate = (BitConverter.ToInt32(val, 0)) * 8 / 1000;

                        // 재생 시간 = 파일크기 / 비트레이트
                        var time = File.Length / bitrate / 125;
                        file_element.lines = bitrate + "kbps" + '\r' + time + "sec";

                        //ele_collection에 instance 추가
                        ele_collection.Add(file_element);
                    }

                    // 나머지 format 파일은 무시한다.
                    else
                    {
                        // 폴더에 진입했는데 (파일은 존재하지만) 적절한 파일이 없을 경우
                        // 저장 위치 ele_collection에서 instance를 제거한다.
                        if (ele_collection.Count > 0)
                        { ele_collection.RemoveAt(ele_collection.Count - 1); }
                    }
                }
                table_flag = 0;
            }
        }

        private void axHwpCtrl1_NotifyMessage(object sender, AxHWPCONTROLLib._DHwpCtrlEvents_NotifyMessageEvent e)
        {
        }
        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
        }

        // crc32 테이블
        private static readonly uint[] defaultTable =
            {
                0x00000000, 0x77073096, 0xEE0E612C, 0x990951BA,
                0x076DC419, 0x706AF48F, 0xE963A535, 0x9E6495A3,
                0x0EDB8832, 0x79DCB8A4, 0xE0D5E91E, 0x97D2D988,
                0x09B64C2B, 0x7EB17CBD, 0xE7B82D07, 0x90BF1D91,
                0x1DB71064, 0x6AB020F2, 0xF3B97148, 0x84BE41DE,
                0x1ADAD47D, 0x6DDDE4EB, 0xF4D4B551, 0x83D385C7,
                0x136C9856, 0x646BA8C0, 0xFD62F97A, 0x8A65C9EC,
                0x14015C4F, 0x63066CD9, 0xFA0F3D63, 0x8D080DF5,
                0x3B6E20C8, 0x4C69105E, 0xD56041E4, 0xA2677172,
                0x3C03E4D1, 0x4B04D447, 0xD20D85FD, 0xA50AB56B,
                0x35B5A8FA, 0x42B2986C, 0xDBBBC9D6, 0xACBCF940,
                0x32D86CE3, 0x45DF5C75, 0xDCD60DCF, 0xABD13D59,
                0x26D930AC, 0x51DE003A, 0xC8D75180, 0xBFD06116,
                0x21B4F4B5, 0x56B3C423, 0xCFBA9599, 0xB8BDA50F,
                0x2802B89E, 0x5F058808, 0xC60CD9B2, 0xB10BE924,
                0x2F6F7C87, 0x58684C11, 0xC1611DAB, 0xB6662D3D,
                0x76DC4190, 0x01DB7106, 0x98D220BC, 0xEFD5102A,
                0x71B18589, 0x06B6B51F, 0x9FBFE4A5, 0xE8B8D433,
                0x7807C9A2, 0x0F00F934, 0x9609A88E, 0xE10E9818,
                0x7F6A0DBB, 0x086D3D2D, 0x91646C97, 0xE6635C01,
                0x6B6B51F4, 0x1C6C6162, 0x856530D8, 0xF262004E,
                0x6C0695ED, 0x1B01A57B, 0x8208F4C1, 0xF50FC457,
                0x65B0D9C6, 0x12B7E950, 0x8BBEB8EA, 0xFCB9887C,
                0x62DD1DDF, 0x15DA2D49, 0x8CD37CF3, 0xFBD44C65,
                0x4DB26158, 0x3AB551CE, 0xA3BC0074, 0xD4BB30E2,
                0x4ADFA541, 0x3DD895D7, 0xA4D1C46D, 0xD3D6F4FB,
                0x4369E96A, 0x346ED9FC, 0xAD678846, 0xDA60B8D0,
                0x44042D73, 0x33031DE5, 0xAA0A4C5F, 0xDD0D7CC9,
                0x5005713C, 0x270241AA, 0xBE0B1010, 0xC90C2086,
                0x5768B525, 0x206F85B3, 0xB966D409, 0xCE61E49F,
                0x5EDEF90E, 0x29D9C998, 0xB0D09822, 0xC7D7A8B4,
                0x59B33D17, 0x2EB40D81, 0xB7BD5C3B, 0xC0BA6CAD,
                0xEDB88320, 0x9ABFB3B6, 0x03B6E20C, 0x74B1D29A,
                0xEAD54739, 0x9DD277AF, 0x04DB2615, 0x73DC1683,
                0xE3630B12, 0x94643B84, 0x0D6D6A3E, 0x7A6A5AA8,
                0xE40ECF0B, 0x9309FF9D, 0x0A00AE27, 0x7D079EB1,
                0xF00F9344, 0x8708A3D2, 0x1E01F268, 0x6906C2FE,
                0xF762575D, 0x806567CB, 0x196C3671, 0x6E6B06E7,
                0xFED41B76, 0x89D32BE0, 0x10DA7A5A, 0x67DD4ACC,
                0xF9B9DF6F, 0x8EBEEFF9, 0x17B7BE43, 0x60B08ED5,
                0xD6D6A3E8, 0xA1D1937E, 0x38D8C2C4, 0x4FDFF252,
                0xD1BB67F1, 0xA6BC5767, 0x3FB506DD, 0x48B2364B,
                0xD80D2BDA, 0xAF0A1B4C, 0x36034AF6, 0x41047A60,
                0xDF60EFC3, 0xA867DF55, 0x316E8EEF, 0x4669BE79,
                0xCB61B38C, 0xBC66831A, 0x256FD2A0, 0x5268E236,
                0xCC0C7795, 0xBB0B4703, 0x220216B9, 0x5505262F,
                0xC5BA3BBE, 0xB2BD0B28, 0x2BB45A92, 0x5CB36A04,
                0xC2D7FFA7, 0xB5D0CF31, 0x2CD99E8B, 0x5BDEAE1D,
                0x9B64C2B0, 0xEC63F226, 0x756AA39C, 0x026D930A,
                0x9C0906A9, 0xEB0E363F, 0x72076785, 0x05005713,
                0x95BF4A82, 0xE2B87A14, 0x7BB12BAE, 0x0CB61B38,
                0x92D28E9B, 0xE5D5BE0D, 0x7CDCEFB7, 0x0BDBDF21,
                0x86D3D2D4, 0xF1D4E242, 0x68DDB3F8, 0x1FDA836E,
                0x81BE16CD, 0xF6B9265B, 0x6FB077E1, 0x18B74777,
                0x88085AE6, 0xFF0F6A70, 0x66063BCA, 0x11010B5C,
                0x8F659EFF, 0xF862AE69, 0x616BFFD3, 0x166CCF45,
                0xA00AE278, 0xD70DD2EE, 0x4E048354, 0x3903B3C2,
                0xA7672661, 0xD06016F7, 0x4969474D, 0x3E6E77DB,
                0xAED16A4A, 0xD9D65ADC, 0x40DF0B66, 0x37D83BF0,
                0xA9BCAE53, 0xDEBB9EC5, 0x47B2CF7F, 0x30B5FFE9,
                0xBDBDF21C, 0xCABAC28A, 0x53B39330, 0x24B4A3A6,
                0xBAD03605, 0xCDD70693, 0x54DE5729, 0x23D967BF,
                0xB3667A2E, 0xC4614AB8, 0x5D681B02, 0x2A6F2B94,
                0xB40BBE37, 0xC30C8EA1, 0x5A05DF1B, 0x2D02EF8D
            };

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void axHwpCtrl1_NotifyMessage_1(object sender, AxHWPCONTROLLib._DHwpCtrlEvents_NotifyMessageEvent e)
        {
        }
    }
}
