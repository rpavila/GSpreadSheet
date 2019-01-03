using System;
using System.Collections.Generic;
using System.Text;

namespace GSpreadSheet
{
    public class ExecutionResult
    {
        public ResultTypes Result { get; set; }
        public string[] Messages { get; set; }
    }

    public class ExecutionResultWithData<T> : ExecutionResult
    {
        public T Data { get; set; }
    }

    public enum ResultTypes
    {
        Success = 0,
        Warning = 1,
        Error = 2
    }
}
