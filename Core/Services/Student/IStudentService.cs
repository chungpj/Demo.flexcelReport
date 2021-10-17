using System.Collections.Generic;


namespace Core.Services.Student
{
    using Core.Models;
    public interface IStudentService : IBaseService
    {
        Student GetById(long? id);
        IEnumerable<Student> GetAll(int page, int pageSize, ref int count);
        long Add(Student entity);
        bool Update(Student entity);
        bool Delete(Student entity);
        IEnumerable<Student> GetAll();
    }
}
