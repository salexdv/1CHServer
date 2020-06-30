using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Windows;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Web;

namespace _1CHServer
{
    [ComVisible(true), Guid("A2E25A38-E4F8-11E0-810A-EDCC4824019B"), ProgId("AddIn.1CHServer")]
    public class Server : AddInLib.IInitDone, AddInLib.ILanguageExtender
    {
        const string c_AddinName = "1CHServer";
        const int BufferSize = 65136;

        #region APImetafile
        public const uint CF_METAFILEPICT = 3;
        public const uint CF_ENHMETAFILE = 14;

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool CloseClipboard();

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetClipboardData(uint format);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool IsClipboardFormatAvailable(uint format);
        #endregion

        #region "IInitDone implementation"

        public Server()
        {                        
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void New()
        {            
        }

        public void Init([MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            BackWork = new BackgroundWorker();            
            BackWork.DoWork += new DoWorkEventHandler(BackWork_DoWork);
            BackWork.WorkerSupportsCancellation = true;
            Data1C.Object1C = pConnection;            
            Marshal.GetIUnknownForObject(Data1C.Object1C);            
        }

        public void Done()
        {
            Marshal.Release(Marshal.GetIDispatchForObject(Data1C.Object1C));
            Marshal.ReleaseComObject(Data1C.Object1C);
            Data1C.Object1C = null;            
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void GetInfo(ref object[] pInfo)
        {            
            ((System.Array)pInfo).SetValue("2000", 0);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        #region IAsyncEvent
        public void SetEventBufferDepth(int lDepth)
        {
        }

        public void GetEventBufferDepth(int plDepth)
        {
            
        }

        public void ExternalEvent(string bstrSource, string bstrMessage, string bstrData)
        {

        }

        public void CleanBuffer()
        {
        }
        #endregion

        #region "Переменные"
        public TcpListener Listener;
        public BackgroundWorker BackWork;
        public int ListenerPort;
        public string LastQuery;
        public string Result;
        public int Timeout;
        public bool WaitResult;
        #endregion

        public void RegisterExtensionAs(ref string bstrExtensionName)
        {
            bstrExtensionName = c_AddinName;
        }

        #region "Свойства"
        enum Props
        {   //Числовые идентификаторы свойств нашей внешней компоненты
            Port = 0,            
            LastProp = 1
        }

        public void GetNProps(ref int plProps)
        {	//Здесь 1С получает количество доступных из ВК свойств
            plProps = (int)Props.LastProp;
        }

        public void FindProp(string bstrPropName, ref int plPropNum)
        {	//Здесь 1С ищет числовой идентификатор свойства по его текстовому имени
            switch (bstrPropName)
            {
                case "Port":
                case "Порт":
                    plPropNum = (int)Props.Port;
                    break;                
                default:
                    plPropNum = -1;
                    break;
            }
        }

        public void GetPropName(int lPropNum, int lPropAlias, ref string pbstrPropName)
        {	//Здесь 1С (теоретически) узнает имя свойства по его идентификатору. lPropAlias - номер псевдонима
            pbstrPropName = "";
        }

        public void GetPropVal(int lPropNum, ref object pvarPropVal)
        {	//Здесь 1С узнает значения свойств 
            pvarPropVal = null;
            switch (lPropNum)
            {
                case (int)Props.Port:
                    pvarPropVal = ListenerPort.ToString();
                    break;                
            }
        }

        public void SetPropVal(int lPropNum, ref object varPropVal)
        {	//Здесь 1С изменяет значения свойств            
        }

        public void IsPropReadable(int lPropNum, ref bool pboolPropRead)
        {	//Здесь 1С узнает, какие свойства доступны для чтения
            pboolPropRead = true; // Все свойства доступны для чтения
        }

        public void IsPropWritable(int lPropNum, ref bool pboolPropWrite)
        {	//Здесь 1С узнает, какие свойства доступны для записи
            pboolPropWrite = false;            
        }
        #endregion

        #region "Методы"
        enum Methods
        {	//Числовые идентификаторы методов (процедур или функций) нашей внешней компоненты
            Start = 0,
            Stop = 1,
            ReturnResult = 2,
            LastMethod = 3
        }

        public void GetNMethods(ref int plMethods)
        {	//Здесь 1С получает количество доступных из ВК методов
            plMethods = (int)Methods.LastMethod;
        }

        public void FindMethod(string bstrMethodName, ref int plMethodNum)
        {	//Здесь 1С получает числовой идентификатор метода (процедуры или функции) по имени (названию) процедуры или функции
            plMethodNum = -1;
            switch (bstrMethodName)
            {
                case "Start":
                case "Запустить":
                    plMethodNum = (int)Methods.Start;
                    break;
                case "Stop":
                case "Остановить":
                    plMethodNum = (int)Methods.Stop;
                    break;
                case "ReturnResult":
                case "ВернутьРезультат":
                    plMethodNum = (int)Methods.ReturnResult;
                    break;                
            }
        }

        public void GetMethodName(int lMethodNum, int lMethodAlias, ref string pbstrMethodName)
        {	//Здесь 1С (теоретически) получает имя метода по его идентификатору. lMethodAlias - номер синонима.
            pbstrMethodName = "";
        }

        public void GetNParams(int lMethodNum, ref int plParams)
        {	//Здесь 1С получает количество параметров у метода (процедуры или функции)
            switch (lMethodNum)
            {
                case (int)Methods.Start:
                    plParams = 1;
                    break;
                case (int)Methods.Stop:
                    plParams = 0;
                    break;
                case (int)Methods.ReturnResult:
                    plParams = 1;
                    break;                
            }
        }

        public void GetParamDefValue(int lMethodNum, int lParamNum, ref object pvarParamDefValue)
        {	//Здесь 1С получает значения параметров процедуры или функции по умолчанию            
            pvarParamDefValue = null;            
        }

        public void HasRetVal(int lMethodNum, ref bool pboolRetValue)
        {	//Здесь 1С узнает, возвращает ли метод значение (т.е. является процедурой или функцией)
            pboolRetValue = true;  //Все методы у нас будут функциями (т.е. будут возвращать значение). 
        }

        public void CallAsProc(int lMethodNum, ref System.Array paParams)
        {	//Здесь внешняя компонента выполняет код процедур. А процедур у нас нет.
        }

        public void CallAsFunc(int lMethodNum, ref object pvarRetValue, ref System.Array paParams)
        {	//Здесь внешняя компонента выполняет код функций.			
            pvarRetValue = 0; //Возвращаемое значение метода для 1С

            switch (lMethodNum) //Порядковый номер метода
            {
                // Метод "Запустить"
                #region Запустить
                case (int)Methods.Start:
                    {
                        
                        try
                        {                            
                            // Получаем порт, на котором будет работать сервер                            
                            ListenerPort = Convert.ToInt32(paParams.GetValue(0));
                                                                                   
                            if (Listener == null)
                            {
                                Log("Начало запуска сервера (" + ListenerPort.ToString() + ")");
                                Listener = new TcpListener(IPAddress.Any, ListenerPort);
                                Listener.Start();
                                BackWork.RunWorkerAsync();
                            }
                            else
                                Log("Сервер уже запущен");
                        }
                        catch (Exception E)
                        {
                            CreateException(E.Message);
                        }                        
                        break;
                    }
                #endregion
                // конец метода "Запустить"					

                // Метод "Остановить"
                #region Остановить
                case (int)Methods.Stop:
                    {                                                
                        try
                        {
                            if (Listener != null)
                            {                                
                                BackWork.CancelAsync();
                                Listener.Stop();
                                WaitResult = false;
                                Listener = null;
                                Log("Сервер остановлен");
                            }
                            else
                                Log("Сервер не запущен");
                        }
                        catch (Exception E)
                        {
                            CreateException(E.Message);
                        }
                        break;
                    }
                #endregion

                // Метод "ВернутьРезультат"
                #region ВернутьРезультат
                case (int)Methods.ReturnResult:
                    {                        
                        Result = Convert.ToString(paParams.GetValue(0));                        
                        WaitResult = false;                        
                        break;
                    }
                #endregion
                // конец метода "ВернутьРезультат"			                
            }
        }
        #endregion

        #region ДополнительныеФункции
        // Процедура выполняет запуск указанного кода в 1С Предприятии
        //
        public void Log(string Message)
        {
            try
            {                
                Data1C.AsyncEvent.GetType().InvokeMember("ExternalEvent", System.Reflection.BindingFlags.InvokeMethod, null, Data1C.AsyncEvent, new object[3] { "1CHServer", "Log", DateTime.Now.ToString() + " - " + Message });
                System.Threading.Thread.Sleep(50);
            }
            catch
            {                
            }
        }

        // Процедура генерирует для 1С внешнее событие с указанным кодом в данных
        //
        public void Run(string Code)
        {                                    
            try
            {                                    
                Data1C.AsyncEvent.GetType().InvokeMember("ExternalEvent", System.Reflection.BindingFlags.InvokeMethod, null, Data1C.AsyncEvent, new object[3] { "1CHServer", "Execute", Code });
            }
            catch
            {                
            }                            
        }

        public void CreateException(string Text)
        {                        
            System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();            
            ExcInfo.wCode = 1006; //Вид пиктограммы            
            ExcInfo.bstrDescription = Text;            
            ExcInfo.bstrSource = c_AddinName;

            Data1C.ErrorLog.AddError("", ref ExcInfo);            
            throw new COMException("An exception has occurred.");            
        }
        

        // Прослушивание порта в отдельном потоке
        //
        private void BackWork_DoWork(object sender, DoWorkEventArgs e)
        {
            Log("Сервер успешно запущен");
                  
            while (true)
            {
                Result = "";

                byte[] responseBuffer = null;

                TcpClient Client = null;
                try
                {
                    // Запускаем прослушивание порта                    
                    Client = Listener.AcceptTcpClient();
                    Log("Входящее соединение (" + Client.Client.RemoteEndPoint.ToString() + ")");
                }
                catch
                {
                    GC.Collect();
                    GC.WaitForFullGCComplete();
                    break;
                }

                // Входящее соединение. Получаем поток клиента
                NetworkStream ClientStream = Client.GetStream();

                string response = String.Empty;

                // Получаем запрос из потока                    
                try
                {
                    byte[] buffer = new byte[BufferSize];
                    int ReadBytes = 0;
                    ReadBytes = ClientStream.Read(buffer, 0, BufferSize);
                    string request = Encoding.UTF8.GetString(buffer).Substring(0, ReadBytes);
                    MessageBox.Show(request);

                    string[] stringsRequest = request.Split('/');

                    string Query = stringsRequest[1];
                    MessageBox.Show(Query);
                    Query = Query.Substring(0, Query.Length - 4);
                    MessageBox.Show(Query);
                    Query = Query.TrimEnd();
                    MessageBox.Show(Query);
                    Query = HttpUtility.UrlDecode(Query);
                    Query = Query.Replace("&&", "+");
                    MessageBox.Show(Query);
                    LastQuery = Query;
                    Log("Текст запроса: \r\n" + Query);

                    WaitResult = true;
                    Run(Query);

                    while (WaitResult)
                    {
                        System.Threading.Thread.Sleep(50);
                    }

                    // Возвращаем клиенту ответ                                
                    if (String.IsNullOrEmpty(Result.Trim()))
                        Log("Возвращен пустой ответ");
                    else
                        Log("Возвращен ответ: " + Result);
                    while (Result.Length < 6)
                    {
                        Result += " ";
                    }

                    responseBuffer = Encoding.UTF8.GetBytes(Result);
                    ClientStream.Write(responseBuffer, 0, responseBuffer.Length);
                    ClientStream.Close();

                    Query = null;
                    stringsRequest = null;
                    buffer = null;

                    responseBuffer = null;
                    Result = null;
                    Client = null;
                    ClientStream.Dispose();                    
                }
                catch (Exception Excep)
                {
                    CreateException("Error: " + Excep.Message);
                }

                GC.Collect();
                GC.WaitForFullGCComplete();
            }
        }
        #endregion      
    }
}
