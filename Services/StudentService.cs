using CreateExcelSheet.Data;

namespace CreateExcelSheet.Services
{
    public class StudentService : IStudentService
    {
        public List<Students> Students()
        {
            List<Students> StudentsList = new()
            {
                new Students() { Id = 1, Name = "Ade Ola", Ca = 23, Exam = 45 },
                new Students() { Id = 2, Name = "Hassan Usman", Ca = 25, Exam = 55 },
                new Students() { Id = 3, Name = "Emeka Ngozi", Ca = 39, Exam = 40 },
                new Students() { Id = 4, Name = "Hunsu Mauyon", Ca = 30, Exam = 56 },
                new Students() { Id = 5, Name = "Ojo Balikis", Ca = 23, Exam = 50 },
                new Students() { Id = 6, Name = "Tambuwal Yusuf", Ca = 34, Exam = 41 },
                new Students() { Id = 7, Name = "Ohiakim Ebuka", Ca = 27, Exam = 47 },
                new Students() { Id = 8, Name = "Sohe Sede", Ca = 31, Exam = 52 },
                new Students() { Id = 9, Name = "Kunle Lola", Ca = 30, Exam = 48 },
                new Students() { Id = 10, Name = "Aliyu Aliyat", Ca = 25, Exam = 58 },
            };
            return StudentsList.OrderBy(x => x.Name).ToList();
        }
    }
}
