using CreateExcelSheet.Services;
using Microsoft.OpenApi.Models;

namespace CreateExcelSheet
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            builder.Services.AddControllers();

            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();
            builder.Services.AddTransient<IStudentService, StudentService>();
            builder.Services.AddTransient<IExcelGenerationService, ExcelGenerationService>();
            builder.Services.AddSwaggerGen(options =>
            {
                options.SwaggerDoc("v1", new OpenApiInfo
                {
                    Title = "CreateExcelSheet",
                    Version = "v1"
                });
            });

            var app = builder.Build();

            app.UseSwagger();
            app.UseSwaggerUI(c => c.SwaggerEndpoint("/swagger/v1/swagger.json", "CreateExcelSheet v1"));

            app.UseHttpsRedirection();
            app.MapControllers();

            app.Run();
        }
    }
}