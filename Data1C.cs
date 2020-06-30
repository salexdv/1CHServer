using System;
using System.Collections.Generic;
using System.Text;

namespace _1CHServer
{
    class Data1C
    {
        public static object Object1C
        {
            get
            {
                return m_Object1C;
            }
            set
            {
                m_Object1C = value;
                m_AsyncEvent = (AddInLib.IAsyncEvent)value;
                m_ErrorInfo = (AddInLib.IErrorLog)value;

            }
        }
        
        public static AddInLib.IAsyncEvent AsyncEvent
        {
            get
            {
                return m_AsyncEvent;
            }
        }

        public static AddInLib.IErrorLog ErrorLog
        {
            get
            {
                return m_ErrorInfo;
            }
        }


        private static object m_Object1C;
        private static AddInLib.IAsyncEvent m_AsyncEvent;
        private static AddInLib.IErrorLog m_ErrorInfo;
    }
}
