using CreateExcelSheet.Data;

namespace CreateExcelSheet.Services
{
    public interface IExcelGenerationService
    {
        Task<MemoryStream> GenerateStudentList(List<Students> list);
    }
}
