using excelTask3.Interfaces;

namespace excelTask3.Service
{
    public class ModelService
    {
        protected readonly IExcelProcess _excelProcess;
        
        public ModelService(IExcelProcess excelProcess)
        {
            _excelProcess = excelProcess;
        }
    }
}